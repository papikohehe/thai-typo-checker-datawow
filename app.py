import streamlit as st
import docx
import thaispellcheck
import html as html_lib
import re
from pythainlp.spell import spell
from pythainlp.tokenize import word_tokenize

# Constants
PHINTHU = "\u0E3A"

VALID_PERIOD_PATTERNS = [
    r"\b[0-9]+\.",                  # Arabic numeral lists: 1., 2.
    r"\b[ก-ฮ]\.",                   # Thai alphabetical lists: ก., ข.
    r"\b[๐-๙]+\.",                 # Thai numeral lists: ๒., ๓.
    r"\b[๐-๙]{1,2}\.[๐-๙]{1,2}",   # Thai time: ๑๐.๑๐
    r"\bพ\.ศ\.",                   # พ.ศ.
    r"\bค\.ศ\.",                   # ค.ศ.
    r"\.{3,}"                       # Ellipses: ..., ..........
]

# UI
st.title("Thai Spellchecker for DOCX")
st.write("🔍 Upload a `.docx` file to find and highlight:")
st.markdown("""
- 🔴 **High Error** (❗ Found by both spellcheckers)<br>
- 🟠 **Error** (❗ Found by one spellchecker)<br>
- 🟧 **Phinthu** (◌ฺ character)<br>
- 🟣 **Apostrophe** `'`<br>
- 🔵 **Invalid Period** `.`
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Choose a Word document", type="docx")

filters = st.multiselect(
    "Filter by error type:",
    ["High Error", "Error", "Phinthu (◌ฺ)", "Apostrophe (`')", "Invalid Period"],
    default=["High Error", "Error", "Phinthu (◌ฺ)", "Apostrophe (`')", "Invalid Period"]
)


def find_invalid_periods(text):
    invalid_indices = []
    for match in re.finditer(r"\.", text):
        is_valid = False
        for pattern in VALID_PERIOD_PATTERNS:
            context = text[max(0, match.start() - 5):match.end() + 5]
            if re.search(pattern, context):
                is_valid = True
                break
        if not is_valid:
            invalid_indices.append(match.start())
    return invalid_indices


def highlight_invalid_periods(text, invalid_indices):
    offset = 0
    for idx in invalid_indices:
        real_idx = idx + offset
        text = text[:real_idx] + "<mark style='background-color:#add8e6;'>.</mark>" + text[real_idx + 1:]
        offset += len("<mark style='background-color:#add8e6;'>.</mark>") - 1
    return text


def safe_check(text):
    try:
        marked = thaispellcheck.check(text, autocorrect=False)
        if len(marked.replace("<คำผิด>", "").replace("</คำผิด>", "")) < len(text) - 5:
            return text
        return marked
    except Exception:
        return text


def cross_check_spelling(text):
    tokens = word_tokenize(text)
    marked_text = thaispellcheck.check(text, autocorrect=False)
    thaispell_errors = re.findall(r"<คำผิด>(.*?)</คำผิด>", marked_text)
    pythai_errors = [word for word in tokens if word not in spell(word)]

    all_errors = set(thaispell_errors + pythai_errors)
    high_errors = []
    errors = []

    for word in all_errors:
        in_thaispell = word in thaispell_errors
        in_pythai = word in pythai_errors
        if in_thaispell and in_pythai:
            high_errors.append(word)
        else:
            errors.append(word)

    return high_errors, errors


def check_docx(file):
    doc = docx.Document(file)
    paragraphs = doc.paragraphs
    total = len(paragraphs)
    results = []

    progress_bar = st.progress(0, text="Processing...")

    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue

        has_phinthu = PHINTHU in text
        has_apostrophe = "'" in text
        invalid_periods = find_invalid_periods(text)
        marked = safe_check(text)
        high_errors, errors = cross_check_spelling(text)

        if "<คำผิด>" in marked or has_phinthu or has_apostrophe or invalid_periods or high_errors or errors:
            results.append({
                "line_no": i + 1,
                "original": text,
                "marked": marked,
                "has_phinthu": has_phinthu,
                "has_apostrophe": has_apostrophe,
                "invalid_periods": invalid_periods,
                "high_errors": high_errors,
                "errors": errors
            })

        progress = int((i + 1) / total * 100)
        progress_bar.progress(progress, text=f"Processing paragraph {i + 1} of {total} ({progress}%)")

    progress_bar.empty()
    return results


def render_html(results, filters):
    html = "<style> mark { padding: 2px 4px; border-radius: 3px; } </style>"
    for item in results:
        line_no = item["line_no"]
        original = html_lib.escape(item["original"])
        marked = html_lib.escape(item["marked"])
        has_phinthu = item["has_phinthu"]
        has_apostrophe = item["has_apostrophe"]
        invalid_periods = item["invalid_periods"]
        high_errors = item["high_errors"]
        errors = item["errors"]

        should_display = (
            ("High Error" in filters and high_errors)
            or ("Error" in filters and errors)
            or ("Phinthu (◌ฺ)" in filters and has_phinthu)
            or ("Apostrophe (`')" in filters and has_apostrophe)
            or ("Invalid Period" in filters and invalid_periods)
        )

        if not should_display:
            continue

        marked = marked.replace("&lt;คำผิด&gt;", "<mark style='background-color:#ffcccc;'>")
        marked = marked.replace("&lt;/คำผิด&gt;", "</mark>")
        marked = marked.replace(PHINTHU, "<mark style='background-color:#ffb84d;'>◌ฺ</mark>")

        def highlight_apostrophes(text):
            def replacer(match):
                content = match.group(1)
                return ">" + content.replace("'", "<mark style='background-color:#d5b3ff;'>'</mark>") + "<"
            return re.sub(r">(.*?)<", replacer, text)

        marked = highlight_apostrophes(marked)
        marked = highlight_invalid_periods(marked, invalid_periods)

        html += f"<div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;'>"
        html += f"<b>🔎 Line {line_no}</b><br>"

        if high_errors and "High Error" in filters:
            html += f"<span style='color:#cc0000;'>🔴 High Error: {', '.join(high_errors)}</span><br>"

        if errors and "Error" in filters:
            html += f"<span style='color:#ff6600;'>🟠 Error: {', '.join(errors)}</span><br>"

        if has_phinthu and "Phinthu (◌ฺ)" in filters:
            html += f"<span style='color:#d2691e;'>🟧 Found unexpected dot (◌ฺ)</span><br>"

        if has_apostrophe and "Apostrophe (`')" in filters:
            html += f"<span style='color:#800080;'>🟣 Found apostrophe `'`</span><br>"

        if invalid_periods and "Invalid Period" in filters:
            html += f"<span style='color:#0055aa;'>🔵 Found suspicious period `.` usage</span><br>"

        html += f"<code style='color:gray;'>{original}</code><br>"
        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{marked}</div></div>"

    return html


# Main app logic
if uploaded_file:
    with st.spinner("🔎 Checking for typos and issues..."):
        results = check_docx(uploaded_file)
        if results:
            st.markdown(render_html(results, filters), unsafe_allow_html=True)
        else:
            st.success("✅ No typos, apostrophes, ◌ฺ characters, or invalid periods found!")
