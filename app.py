import streamlit as st
import docx
import thaispellcheck
import html as html_lib
import re
from pythainlp.spell import spell

# Constants
PHINTHU = "\u0E3A"

VALID_PERIOD_PATTERNS = [
    r"\b[0-9]+\.",                  # Arabic numeral lists: 1., 2.
    r"\b[ก-ฮ]\.",                   # Thai alphabetical lists: ก., ข.
    r"\b[๐-๙]+\.",                 # Thai numeral lists: ๒., ๓.
    r"\b[๐-๙]{1,2}\.[๐-๙]{1,2}",   # Thai time: ๑๐.๑๐
    r"\bพ\.ศ\.",                   # พ.ศ.
    r"\bค\.ศ\.",                   # ค.ศ.
    r"\.{3,}"                       # Ellipses
]

# UI
st.title("Thai Spellchecker for DOCX")
st.write("🔍 Upload a `.docx` file to find and highlight issues.")
st.markdown("""
- ❌ Thai spelling errors (🔴 red)<br>
- ⚠️ Unexpected Thai dot ◌ฺ (🟠 orange)<br>
- ⚠️ Misused apostrophes `'` (🟣 purple)<br>
- ⚠️ Invalid period use `.` (🔵 blue)
""", unsafe_allow_html=True)

engine_choice = st.selectbox("Select Spellcheck Engine", ["thaispellcheck", "pythainlp"])

error_filters = st.multiselect(
    "Filter error types",
    ["Spelling", "Phinthu", "Apostrophe", "Invalid Period"],
    default=["Spelling", "Phinthu", "Apostrophe", "Invalid Period"]
)

uploaded_file = st.file_uploader("Choose a Word document", type="docx")


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


def safe_check(text, engine="thaispellcheck"):
    try:
        if engine == "thaispellcheck":
            marked = thaispellcheck.check(text, autocorrect=False)
            if len(marked.replace("<คำผิด>", "").replace("</คำผิด>", "")) < len(text) - 5:
                return text  # fallback
            return marked
        elif engine == "pythainlp":
            words = text.split()
            mistakes = set(spell(text))
            marked = ""
            for word in words:
                if word in mistakes:
                    marked += f"<คำผิด>{word}</คำผิด> "
                else:
                    marked += word + " "
            return marked.strip()
    except Exception:
        return text


def check_docx(file, engine="thaispellcheck"):
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
        marked = safe_check(text, engine)

        if "<คำผิด>" in marked or has_phinthu or has_apostrophe or invalid_periods:
            results.append({
                "line_no": i + 1,
                "original": text,
                "marked": marked,
                "has_phinthu": has_phinthu,
                "has_apostrophe": has_apostrophe,
                "invalid_periods": invalid_periods
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

        # Filter logic
        skip = True
        if "Spelling" in filters and "<คำผิด>" in item["marked"]:
            skip = False
        if "Phinthu" in filters and has_phinthu:
            skip = False
        if "Apostrophe" in filters and has_apostrophe:
            skip = False
        if "Invalid Period" in filters and invalid_periods:
            skip = False
        if skip:
            continue

        # Highlight
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
        html += f"<b>❌ Line {line_no}</b><br>"

        if has_phinthu and "Phinthu" in filters:
            html += f"<span style='color:#d00;'>⚠️ Found unexpected dot (◌ฺ) — possibly OCR or typing error.</span><br>"

        if has_apostrophe and "Apostrophe" in filters:
            html += f"<span style='color:#800080;'>⚠️ Found apostrophe `'` — may be unintended.</span><br>"

        if invalid_periods and "Invalid Period" in filters:
            html += f"<span style='color:#0055aa;'>⚠️ Found suspicious period `.` usage — not in พ.ศ., ค.ศ., list formats, Thai time, or ellipses.</span><br>"

        html += f"<code style='color:gray;'>{original}</code><br>"
        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{marked}</div></div>"

    return html


# Main logic
if uploaded_file:
    with st.spinner("🔎 Checking for typos and issues..."):
        results = check_docx(uploaded_file, engine=engine_choice)
        if results:
            st.markdown(render_html(results, error_filters), unsafe_allow_html=True)
        else:
            st.success("✅ No typos, apostrophes, ◌ฺ characters, or invalid periods found!")
