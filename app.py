import streamlit as st

import docx

import thaispellcheck

import html as html_lib

import re



# Constants

PHINTHU = "\u0E3A"

COMMON_ERRORS = { "เข่น", "ล่ง", "สาย", "ขี้", "ขื่อ", "ศักดิ๋", "ขัก", "ฃ้ือ", "ชื้อ", "แกไข", "ที'",

    "บาย", "ข่วย", "แก่ไข", "สมาซิก", "ไมได้", "ครังที", "ฤทธ๋", "ศักด๋", "ด้งนี้",

    "มดิ", "ซัดเจน", "เพิ่มเดิม", "เลียหาย", "ส่ง", "มบุษยชน", "สิทธิ๔", "เดิมฺ",

    "ขุม", "นันทํ", "ๆ", "ไซด์", "เร้ยีน", "ประจา", "ที", "สา", "คู", "ชอง", "ทนี่ง",

    "เหลีอมลา", "ลี", "ซาน", "โช๊ะ", "โฃ๊ะ", "สถาน", "เมือ", "กัมพูขา", "สิทธิมบุษยชน",

    "ศคินันท์", "กณวีร์", "๙0", "ชั้น", "ลูก", "ศักดิ์", "ทันตแพทย์สภา", "แกไข", "ไว",

    "รับพิง", "คิริโรจน์", "ชักถาม" }



# Valid patterns for Thai period usage

VALID_PERIOD_PATTERNS = [

    r"\b[0-9]+\.",                  # Arabic numeral lists: 1., 2.

    r"\b[ก-ฮ]\.",                   # Thai alphabetical lists: ก., ข.

    r"\b[๐-๙]+\.",                 # Thai numeral lists: ๒., ๓.

    r"\b[๐-๙]{1,2}\.[๐-๙]{1,2}",   # Thai time: ๑๐.๑๐

    r"\bพ\.ศ\.",                   # พ.ศ.

    r"\bค\.ศ\.",                   # ค.ศ.

    r"\.{3,}"                      # Ellipses: ..., ..........

]



# UI

st.title("Thai Spellchecker for DOCX")

st.write("🔍 Upload a `.docx` file to find and highlight:")

st.markdown("""

- ❌ Thai spelling errors (🔴 red)<br>

- ⚠️ Unexpected Thai dot ◌ฺ (🟠 orange)<br>

- ⚠️ Misused apostrophes `'` (🟣 purple)<br>

- ⚠️ Invalid period use `.` (🔵 blue)<br>

- ⚠️ Common error words (🟡 yellow)

""", unsafe_allow_html=True)



uploaded_file = st.file_uploader("Choose a Word document", type="docx")





# Helper functions

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





def find_common_errors(text):

    return [word for word in COMMON_ERRORS if word in text]





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

        common_errors = find_common_errors(text)

        marked = safe_check(text)



        if "<คำผิด>" in marked or has_phinthu or has_apostrophe or invalid_periods or common_errors:

            results.append({

                "line_no": i + 1,

                "original": text,

                "marked": marked,

                "has_phinthu": has_phinthu,

                "has_apostrophe": has_apostrophe,

                "invalid_periods": invalid_periods,

                "common_errors": common_errors

            })



        progress = int((i + 1) / total * 100)

        progress_bar.progress(progress, text=f"Processing paragraph {i + 1} of {total} ({progress}%)")



    progress_bar.empty()

    return results





def render_html(results):

    def escape(text): return html_lib.escape(text)



    def mark(text, color):

        return f"<mark style='background-color:{color};'>{escape(text)}</mark>"



    html = "<style> mark { padding: 2px 4px; border-radius: 3px; } </style>"



    for item in results:

        line_no = item["line_no"]

        original = escape(item["original"])

        raw_text = item["marked"]



        # Step 1: Replace <คำผิด> tags with safe placeholders

        raw_text = raw_text.replace("<คำผิด>", "[[WRONG_OPEN]]").replace("</คำผิด>", "[[WRONG_CLOSE]]")



        # Step 2: Escape the entire text (placeholders preserved)

        safe_text = escape(raw_text)



        # Step 3: Re-insert highlight marks for spellcheck errors

        safe_text = safe_text.replace("[[WRONG_OPEN]]", "<mark style='background-color:#ffcccc;'>")

        safe_text = safe_text.replace("[[WRONG_CLOSE]]", "</mark>")



        # Step 4: Highlight ◌ฺ

        safe_text = safe_text.replace(escape(PHINTHU), mark(PHINTHU, "#ffb84d"))



        # Step 5: Highlight apostrophes

        safe_text = re.sub(

            r"(>[^<]*)'([^<]*<)",

            lambda m: f"{m.group(1)}<mark style='background-color:#d5b3ff;'>'</mark>{m.group(2)}",

            safe_text

        )



        # Step 6: Highlight invalid periods

        safe_text = re.sub(

            r"(?<!\w)(\.)(?!\w)",

            lambda m: mark(".", "#add8e6"),

            safe_text

        )



        # Step 7: Highlight common errors

        for error_word in COMMON_ERRORS:

            safe_text = safe_text.replace(

                escape(error_word),

                mark(error_word, "#ffff66")  # Yellow

            )



        # Final output block

        html += f"<div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;'>"

        html += f"<b>❌ Line {line_no}</b><br>"



        if item["has_phinthu"]:

            html += f"<span style='color:#d00;'>⚠️ Found unexpected dot (◌ฺ)</span><br>"



        if item["has_apostrophe"]:

            html += f"<span style='color:#800080;'>⚠️ Found apostrophe `'`</span><br>"



        if item["invalid_periods"]:

            html += f"<span style='color:#0055aa;'>⚠️ Found suspicious period `.`</span><br>"



        if item.get("common_errors"):

            html += f"<span style='color:#b58900;'>⚠️ Found common error words: {', '.join(item['common_errors'])}</span><br>"



        html += f"<code style='color:gray;'>{original}</code><br>"

        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{safe_text}</div></div>"



    return html





# Main app logic

if uploaded_file:

    with st.spinner("🔎 Checking for typos and issues..."):

        results = check_docx(uploaded_file)

        if results:

            try:

                st.markdown(render_html(results), unsafe_allow_html=True)

            except Exception as e:

                st.error("🚨 Error rendering HTML.")

                st.exception(e)

        else:

            st.success("✅ No typos, apostrophes, ◌ฺ characters, invalid periods, or common errors found!")

