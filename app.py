import streamlit as st
import docx
import thaispellcheck
import html as html_lib
import re

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

try:
    html_output = render_html(results)
    if "<" in html_output and "<ma<mark" in html_output:
        st.warning("⚠️ HTML structure issue detected. A <mark> may be inserted into a tag.")
    st.markdown(html_output, unsafe_allow_html=True)
except Exception as e:
    st.error("HTML rendering failed.")
    st.exception(e)

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


def render_html(results):
    from html import escape

    def mark(text, color):
        return f"<mark style='background-color:{color};'>{escape(text)}</mark>"

    html = "<style> mark { padding: 2px 4px; border-radius: 3px; } </style>"

    for item in results:
        line_no = item["line_no"]
        original = escape(item["original"])
        raw_text = item["marked"]

        # Step 1: Safely highlight <คำผิด>
        raw_text = raw_text.replace("<คำผิด>", "<<<WRONG>>>").replace("</คำผิด>", "<<<END>>>")

        # Step 2: Escape everything
        safe_text = escape(raw_text)

        # Step 3: Restore highlight tags
        safe_text = safe_text.replace("<<<WRONG>>>", "<mark style='background-color:#ffcccc;'>")
        safe_text = safe_text.replace("<<<END>>>", "</mark>")

        # Step 4: Highlight ◌ฺ
        safe_text = safe_text.replace(escape(PHINTHU), mark(PHINTHU, "#ffb84d"))

        # Step 5: Highlight apostrophes only between tags
        safe_text = re.sub(
            r"(>[^<]*)'([^<]*<)",
            lambda m: f"{m.group(1)}<mark style='background-color:#d5b3ff;'>'</mark>{m.group(2)}",
            safe_text
        )

        # Step 6: Highlight invalid periods using their indexes
        # (We must work with already-marked HTML, so we skip index-based insertion — instead, match isolated dots)
        safe_text = re.sub(
            r"(?<!\w)(\.)(?!\w)",
            lambda m: mark(".", "#add8e6"),
            safe_text
        )

        # Step 7: Wrap output
        html += f"<div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;'>"
        html += f"<b>❌ Line {line_no}</b><br>"

        if item["has_phinthu"]:
            html += f"<span style='color:#d00;'>⚠️ Found unexpected dot (◌ฺ)</span><br>"

        if item["has_apostrophe"]:
            html += f"<span style='color:#800080;'>⚠️ Found apostrophe `'`</span><br>"

        if item["invalid_periods"]:
            html += f"<span style='color:#0055aa;'>⚠️ Found suspicious period `.`</span><br>"

        html += f"<code style='color:gray;'>{original}</code><br>"
        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{safe_text}</div></div>"

    return html
# Main app logic
if uploaded_file:
    with st.spinner("🔎 Checking for typos and issues..."):
        results = check_docx(uploaded_file)
        if results:
            st.markdown(render_html(results), unsafe_allow_html=True)
        else:
            st.success("✅ No typos, apostrophes, ◌ฺ characters, or invalid periods found!")




