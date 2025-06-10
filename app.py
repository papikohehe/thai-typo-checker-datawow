import streamlit as st
import docx
import thaispellcheck

PHINTHU = "\u0E3A"

st.title("Thai Spellchecker for DOCX")
st.write("🔍 Upload a `.docx` file to find and highlight:")
st.markdown("""
- ❌ Thai spelling errors (🔴 red)
- ⚠️ Unexpected Thai dot ◌ฺ (🟠 orange)
- ⚠️ Misused apostrophes `'` (🟣 purple)
""")

uploaded_file = st.file_uploader("Choose a Word document", type="docx")

def check_docx(file):
    doc = docx.Document(file)
    results = []

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        has_phinthu = PHINTHU in text
        has_apostrophe = "'" in text

        marked = thaispellcheck.check(text, autocorrect=False)

        if "<คำผิด>" in marked or has_phinthu or has_apostrophe:
            results.append({
                "line_no": i + 1,
                "original": text,
                "marked": marked,
                "has_phinthu": has_phinthu,
                "has_apostrophe": has_apostrophe
            })
    return results

def render_html(results):
    html = "<style> mark { padding: 2px 4px; border-radius: 3px; } </style>"
    for item in results:
        line_no = item["line_no"]
        original = item["original"]
        marked = item["marked"]
        has_phinthu = item["has_phinthu"]
        has_apostrophe = item["has_apostrophe"]

        # Highlight typos (in red)
        marked = marked.replace("<คำผิด>", "<mark style='background-color:#ffcccc;'>")
        marked = marked.replace("</คำผิด>", "</mark>")

        # Highlight phinthu (in orange)
        marked = marked.replace(PHINTHU, "<mark style='background-color:#ffb84d;'>◌ฺ</mark>")

        # Highlight apostrophes (in purple)
        marked = marked.replace("'", "<mark style='background-color:#d5b3ff;'>'</mark>")

        html += f"<div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;'>"
        html += f"<b>❌ Line {line_no}</b><br>"

        if has_phinthu:
            html += f"<span style='color:#d00;'>⚠️ Found unexpected dot (◌ฺ) — possibly OCR or typing error.</span><br>"

        if has_apostrophe:
            html += f"<span style='color:#800080;'>⚠️ Found apostrophe `'` — may be unintended.</span><br>"

        html += f"<code style='color:gray;'>{original}</code><br>"
        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{marked}</div></div>"
    return html

if uploaded_file:
    with st.spinner("🔎 Checking for typos and issues..."):
        results = check_docx(uploaded_file)
        if results:
            st.markdown(render_html(results), unsafe_allow_html=True)
        else:
            st.success("✅ No typos, apostrophes, or ◌ฺ characters found!")
