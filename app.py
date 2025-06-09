import streamlit as st
import docx
import thaispellcheck

st.title("Thai Spellchecker for DOCX")
st.write("🔍 Upload a `.docx` file to find and highlight Thai typos.")

uploaded_file = st.file_uploader("Choose a Word document", type="docx")

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

def render_html(results):
    html = ""
    for line_no, original, marked in results:
        marked = marked.replace("<คำผิด>", "<mark style='background-color:#ffcccc;'>")
        marked = marked.replace("</คำผิด>", "</mark>")
        html += f"<div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;'>"
        html += f"<b>❌ Line {line_no}</b><br>"
        html += f"<code style='color:gray;'>{original}</code><br>"
        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{marked}</div></div>"
    return html

if uploaded_file:
    with st.spinner("🔎 Checking for typos..."):
        results = check_docx(uploaded_file)
        if results:
            st.markdown(render_html(results), unsafe_allow_html=True)
        else:
            st.success("✅ No typos found!")
