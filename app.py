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


@st.cache_data(show_spinner=False)
def cross_check_spelling(text):
    tokens = word_tokenize(text)
    thaispell_errors = set(thaispellcheck.get_errors(text))
    pythainlp_errors = set(w for w in tokens if w not in spell(w))

    high_errors = list(thaispell_errors & pythainlp_errors)
    partial_errors = list((thaispell_errors | pythainlp_errors) - set(high_errors))

    return {"high_errors": high_errors, "errors": partial_errors}


def check_docx(file):
    doc = docx.Document(file)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    results = []

    progress_bar = st.progress(0, text="Processing...")

    for i, text in enumerate(paragraphs):
        has_phinthu = PHINTHU in text
        has_apostrophe = "'" in text
        invalid_periods = find_invalid_periods(text)
        spell_result = cross_check_spelling(text)

        if any([spell_result["high_errors"], spell_result["errors"],
                has_phinthu, has_apostrophe, invalid_periods]):
            # Only now run the slower marking
            marked = safe_check(text)

            results.append({
                "line_no": i + 1,
                "original": text,
                "marked": marked,
                "has_phinthu": has_phinthu,
                "has_apostrophe": has_apostrophe,
                "invalid_periods": invalid_periods,
                "high_errors": spell_result["high_errors"],
                "errors": spell_result["errors"]
            })

        progress = int((i + 1) / len(paragraphs) * 100)
        progress_bar.progress(progress, text=f"Processing paragraph {i + 1} of {len(paragraphs)}")

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
