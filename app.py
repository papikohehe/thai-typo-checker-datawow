import streamlit as st
import docx
import thaispellcheck

st.title("Thai Spellchecker for DOCX")
st.write("🔍 Upload one or more `.docx` files to find and highlight Thai typos.")

uploaded_files = st.file_uploader("Choose Word document(s)", type="docx", accept_multiple_files=True)

def check_docx(file):
    doc = docx.Document(file)
    results = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        marked = thaispellcheck.check(text, autocorrect=False)
        if "<คำผิด>" in marked:
            results.append((i + 1, text, marked))
    return results

def render_html(results, filename):
    html = f"<h3 style='margin-top:2em;'>{filename}</h3>"
    for line_no, original, marked in results:
        marked = marked.replace("<คำผิด>", "<mark style='background-color:#ffcccc;'>")
        marked = marked.replace("</คำผิด>", "</mark>")
        html += f\"\"\"\n        <div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;'>\n            <b>❌ Line {line_no}</b><br>\n            <code style='color:gray;'>{original}</code><br>\n            <div style='margin-top:0.5em;font-size:1.1em;'>{marked}</div>\n        </div>\n        \"\"\"\n    return html

if uploaded_files:
    full_html = \"\"\"
    <style> mark { padding: 2px 4px; border-radius: 3px; } </style>
    \"\"\"\n    for uploaded_file in uploaded_files:
        with st.spinner(f\"🔎 Checking {uploaded_file.name}...\"):\n            results = check_docx(uploaded_file)
            if results:
                full_html += render_html(results, uploaded_file.name)
            else:
                full_html += f\"<h3>{uploaded_file.name}</h3><p>✅ No typos found!</p>\"\n
    st.markdown(full_html, unsafe_allow_html=True)
