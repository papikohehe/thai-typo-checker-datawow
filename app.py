import streamlit as st
import docx
import thaispellcheck
import html as html_lib
import re

PHINTHU = "\u0E3A"
VALID_PERIOD_PATTERNS = [
    r"\b[0-9]+\.",             # Arabic numbered list: 1., 2., etc.
    r"\b[๐-๙]+\.",             # Thai numbered list: ๑., ๒., etc.
    r"\b[ก-ฮ]\.",              # Thai alphabetical list: ก., ข., etc.
    r"\bพ\.ศ\.",               # Buddhist Era
    r"\bค\.ศ\.",               # Christian Era
    r"[๐-๙]{1,2}\.[๐-๙]{1,2}"  # Thai time or decimal number: ๑๒.๓๕
]


st.title("Thai Spellchecker for DOCX (Data Wow)")
st.write("🔍 Upload a `.docx` file to find and highlight:")
st.markdown("""
- ❌ Thai spelling errors (🔴 red)<br>
- ⚠️ Unexpected Thai dot ◌ฺ (🟠 orange)<br>
- ⚠️ Misused apostrophes `'` (🟣 purple)<br>
- ⚠️ Invalid period use `.` (🔵 blue)
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Choose a Word document", type="docx")

def find_invalid_periods(text):
    invalid_indices = []
    for match in re.finditer(r"\.", text):
        is_valid = False
        for pattern in VALID_PERIOD_PATTERNS:
            context = text[max(0, match.start() - 10):match.end() + 10]
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
        text = text[:real_idx] + "<mark style='background-color:#add8e6;'>.</mark>" + text[real_idx+1:]
        offset += len("<mark style='background-color:#add8e6;'>.</mark>") - 1
    return text

def check_docx(file):
    doc = docx.Document(file)
    results = []

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        has_phinthu = PHINTHU in text
        has_apostrophe = "'" in text
        invalid_periods = find_invalid_periods(text)

        marked = thaispellcheck.check(text, autocorrect=False)

        if "<คำผิด>" in marked or has_phinthu or has_apostrophe or invalid_periods:
            results.append({
                "line_no": i + 1,
                "original": text,
                "marked": marked,
                "has_phinthu": has_phinthu,
                "has_apostrophe": has_apostrophe,
                "invalid_periods": invalid_periods
            })
    return results

def render_html(results):
    html = "<style> mark { padding: 2px 4px; border-radius: 3px; } </style>"
    for item in results:
        line_no = item["line_no"]
        original = html_lib.escape(item["original"])
        marked = html_lib.escape(item["marked"])
        has_phinthu = item["has_phinthu"]
        has_apostrophe = item["has_apostrophe"]
        invalid_periods = item["invalid_periods"]

        # Highlight typos (in red)
        marked = marked.replace("&lt;คำผิด&gt;", "<mark style='background-color:#ffcccc;'>")
        marked = marked.replace("&lt;/คำผิด&gt;", "</mark>")

        # Highlight phinthu (in orange)
        marked = marked.replace(PHINTHU, "<mark style='background-color:#ffb84d;'>◌ฺ</mark>")

        # Highlight apostrophes (in purple)
        def highlight_apostrophes(text):
            def replacer(match):
                content = match.group(1)
                return ">" + content.replace("'", "<mark style='background-color:#d5b3ff;'>'</mark>") + "<"
            return re.sub(r">(.*?)<", replacer, text)
        marked = highlight_apostrophes(marked)

        # Highlight invalid periods (in blue)
        marked = highlight_invalid_periods(marked, invalid_periods)

        html += f"<div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;'>"
        html += f"<b>❌ Line {line_no}</b><br>"

        if has_phinthu:
            html += f"<span style='color:#d00;'>⚠️ Found unexpected dot (◌ฺ) — possibly OCR or typing error.</span><br>"

        if has_apostrophe:
            html += f"<span style='color:#800080;'>⚠️ Found apostrophe `'` — may be unintended.</span><br>"

        if invalid_periods:
            html += f"<span style='color:#0055aa;'>⚠️ Found suspicious period `.` usage — not in พ.ศ., ค.ศ., or list formats.</span><br>"

        html += f"<code style='color:gray;'>{original}</code><br>"
        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{marked}</div></div>"
    return html

if uploaded_file:
    with st.spinner("🔎 Checking for typos and issues..."):
        results = check_docx(uploaded_file)
        if results:
            st.markdown(render_html(results), unsafe_allow_html=True)
        else:
            st.success("✅ No typos, apostrophes, ◌ฺ characters, or invalid periods found!")
