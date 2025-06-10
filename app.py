import streamlit as st
import docx
import thaispellcheck
from pythainlp.spell import spell
from pythainlp.tokenize import word_tokenize
import html as html_lib
import re

# Constants
PHINTHU = "\u0E3A"

VALID_PERIOD_PATTERNS = [
    r"\b[0-9]+\.",                  
    r"\b[ก-ฮ]\.",                   
    r"\b[๐-๙]+\.",                 
    r"\b[๐-๙]{1,2}\.[๐-๙]{1,2}",   
    r"\bพ\.ศ\.",                   
    r"\bค\.ศ\.",                   
    r"\.{3,}"                      
]

# UI
st.title("Thai Spellchecker for DOCX (Optimized)")
st.write("🔍 Upload a `.docx` file to highlight issues:")
st.markdown("""
- 🔴 **High Error** (found by both thaispellcheck & pythainlp)<br>
- 🟠 **Error** (found by only one checker)<br>
- 🟡 Unexpected Thai dot ◌ฺ<br>
- 🟣 Misused apostrophes `'`<br>
- 🔵 Invalid period use `.`
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

def cross_check_spelling(text):
    tokens = word_tokenize(text)
    thaispell_errors = set(thaispellcheck.get_errors(text))
    pythainlp_errors = set(w for w in tokens if w not in spell(w))

    high_errors = list(thaispell_errors & pythainlp_errors)
    partial_errors = list((thaispell_errors | pythainlp_errors) - set(high_errors))

    return {"high_errors": high_errors, "errors": partial_errors}

def safe_check(text):
    try:
        marked = thaispellcheck.check(text, autocorrect=False)
        if len(marked.replace("<คำผิด>", "").replace("</คำผิด>", "")) < len(text) - 5:
            return text
        return marked
    except Exception:
        return text

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

def highlight_invalid_periods(text, invalid_indices):
    offset = 0
    for idx in invalid_indices:
        real_idx = idx + offset
        text = text[:real_idx] + "<mark style='background-color:#add8e6;'>.</mark>" + text[real_idx + 1:]
        offset += len("<mark style='background-color:#add8e6;'>.</mark>") - 1
    return text

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

        line_no = item["line_no"]
        original = html_lib.escape(item["original"])
        marked = html_lib.escape(item["marked"])
        marked = marked.replace("&lt;คำผิด&gt;", "<mark style='background-color:#ffa3a3;'>")
        marked = marked.replace("&lt;/คำผิด&gt;", "</mark>")
        marked = marked.replace(PHINTHU, "<mark style='background-color:#ffdb99;'>◌ฺ</mark>")

        def highlight_apostrophes(text):
            return re.sub(r">(.*?)<", lambda m: ">" + m.group(1).replace("'", "<mark style='background-color:#e0c2ff;'>'</mark>") + "<", text)

        marked = highlight_apostrophes(marked)
        marked = highlight_invalid_periods(marked, item["invalid_periods"])

        html += f"<div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;'>"
        html += f"<b>Line {line_no}</b><br>"

        if item["high_errors"]:
            html += f"<span style='color:#d00;'>🔴 High Error: {', '.join(item['high_errors'])}</span><br>"
        if item["errors"]:
            html += f"<span style='color:#e67e00;'>🟠 Error: {', '.join(item['errors'])}</span><br>"
        if item["has_phinthu"]:
            html += f"<span style='color:#ff9900;'>🟡 Unexpected dot (◌ฺ)</span><br>"
        if item["has_apostrophe"]:
            html += f"<span style='color:#800080;'>🟣 Apostrophe `'` found</span><br>"
        if item["invalid_periods"]:
            html += f"<span style='color:#0055aa;'>🔵 Suspicious period `.` usage</span><br>"

        html += f"<code style='color:gray;'>{original}</code><br>"
        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{marked}</div></div>"

    return html

# Main
if uploaded_file:
    with st.spinner("🔎 Checking for typos and issues..."):
        results = check_docx(uploaded_file)
        if results:
            st.markdown(render_html(results, filters), unsafe_allow_html=True)
        else:
            st.success("✅ No typos, apostrophes, ◌ฺ characters, or invalid periods found!")
