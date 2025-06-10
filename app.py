import streamlit as st
import docx
import thaispellcheck
import pythainlp
import html as html_lib
import re
from pythainlp.spell import spell
from pythainlp.tokenize import word_tokenize

# Constants
PHINTHU = "\u0E3A"

# Updated patterns to include Thai numerals and ellipses
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
- ❌ Thai spelling errors (🔴 red)<br>
- ⚠️ Unexpected Thai dot ◌ฺ (🟠 orange)<br>
- ⚠️ Misused apostrophes `'` (🟣 purple)<br>
- ⚠️ Invalid period use `.` (🔵 blue)
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Choose a Word document", type="docx")

filters = st.multiselect(
    "Filter by error type:",
    ["High Error", "Error", "Phinthu (◌ฺ)", "Apostrophe (`')", "Invalid Period"],
    default=["High Error", "Error", "Phinthu (◌ฺ)", "Apostrophe (`')", "Invalid Period"]
)

# Helpers
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
            return text  # fallback if it looks wrong
        return marked
    except Exception:
        return text


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

       spell_result = cross_check_spelling(text)

if spell_result["high_errors"] or spell_result["errors"] or has_phinthu or has_apostrophe or invalid_periods:
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

        progress = int((i + 1) / total * 100)
        progress_bar.progress(progress, text=f"Processing paragraph {i + 1} of {total} ({progress}%)")

    progress_bar.empty()
    return results
    


def render_html(results, filters):
    html = "<style> mark { padding: 2px 4px; border-radius: 3px; } </style>"
    for item in results:
        show = False
        if "High Error" in filters and item["high_errors"]:
            show = True
        if "Error" in filters and item["errors"]:
            show = True
        if "Phinthu (◌ฺ)" in filters and item["has_phinthu"]:
            show = True
        if "Apostrophe (`')" in filters and item["has_apostrophe"]:
            show = True
        if "Invalid Period" in filters and item["invalid_periods"]:
            show = True

        if not show:
            continue

        # Highlight <คำผิด>
        marked = marked.replace("&lt;คำผิด&gt;", "<mark style='background-color:#ffcccc;'>")
        marked = marked.replace("&lt;/คำผิด&gt;", "</mark>")

        # Highlight ◌ฺ
        marked = marked.replace(PHINTHU, "<mark style='background-color:#ffb84d;'>◌ฺ</mark>")

        # Highlight apostrophes
        def highlight_apostrophes(text):
            def replacer(match):
                content = match.group(1)
                return ">" + content.replace("'", "<mark style='background-color:#d5b3ff;'>'</mark>") + "<"
            return re.sub(r">(.*?)<", replacer, text)

        marked = highlight_apostrophes(marked)

        # Highlight invalid periods
        marked = highlight_invalid_periods(marked, invalid_periods)

        html += f"<div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;'>"
        html += f"<b>❌ Line {line_no}</b><br>"

        if has_phinthu:
            html += f"<span style='color:#d00;'>⚠️ Found unexpected dot (◌ฺ) — possibly OCR or typing error.</span><br>"

        if has_apostrophe:
            html += f"<span style='color:#800080;'>⚠️ Found apostrophe `'` — may be unintended.</span><br>"

        if invalid_periods:
            html += f"<span style='color:#0055aa;'>⚠️ Found suspicious period `.` usage — not in พ.ศ., ค.ศ., list formats, Thai time, or ellipses.</span><br>"

        html += f"<code style='color:gray;'>{original}</code><br>"
        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{marked}</div></div>"
    return html

def cross_check_spelling(text):
    results = {
        "high_errors": [],
        "errors": []
    }

    tokens = word_tokenize(text)
    thai_spell_errors = thaispellcheck.get_errors(text)
    pythainlp_errors = []

    for word in tokens:
        if word not in spell(word):
            pythainlp_errors.append(word)

    all_errors = set(thai_spell_errors + pythainlp_errors)

    for word in all_errors:
        in_thaispell = word in thai_spell_errors
        in_pythainlp = word in pythainlp_errors

        if in_thaispell and in_pythainlp:
            results["high_errors"].append(word)
        else:
            results["errors"].append(word)

    return results

# Main app logic
if uploaded_file:
    with st.spinner("🔎 Checking for typos and issues..."):
        results = check_docx(uploaded_file)
       if results:
    st.markdown(render_html(results, filters), unsafe_allow_html=True)
else:
    st.success("✅ No issues found!")
