import streamlit as st
import docx
import thaispellcheck
import html as html_lib
import re

# Constants
PHINTHU = "\u0E3A"
COMMON_ERRORS = {
    "เข่น", "ล่ง", "สาย", "ขี้", "ขื่อ", "ศักดิ๋", "ขัก", "ฃ้ือ", "ชื้อ", "แกไข", "ที'",
    "บาย", "ข่วย", "แก่ไข", "สมาซิก", "ไมได้", "ครังที", "ฤทธ๋", "ศักด๋", "ด้งนี้",
    "มดิ", "ซัดเจน", "เพิ่มเดิม", "เลียหาย", "ส่ง", "มบุษยชน", "สิทธิ๔", "เดิมฺ",
    "ขุม", "นันทํ", "ๆ", "ไซด์", "เร้ยีน", "ประจา", "ที", "สา", "คู", "ชอง", "ทนี่ง",
    "เหลีอมลา", "ลี", "ซาน", "โช๊ะ", "โฃ๊ะ", "สถาน", "เมือ", "กัมพูขา", "สิทธิมบุษยชน",
    "ศคินันท์", "กณวีร์", "๙0", "ชั้น", "ลูก", "ศักดิ์", "ทันตแพทย์สภา", "แกไข", "ไว",
    "รับพิง", "คิริโรจน์", "ชักถาม"
}

# Regex pattern adapted from your Google Sheets formula
REGEX_ERROR_PATTERN = re.compile(r"""(^ | $|([ๆ\)]|ฯลฯ)\S|\S(\(|ฯลฯ)|[ก-ูเ-์][A-Za-z0-9]|[A-Za-z0-9][ก-ูเ-์]|[ฯะาำเ-ๆ][ั-ูๅ็-์]|[ฯะเ-ๆ]ะ|[็-์][ิ-ู็-์]|[เ-ไ]{2,}|[ั-ู]{2,}|[เ-ไ][ก-ฮ]์|[โ-ไ][ก-ฮ]็|[ก-ฮ][็์][ะาำ]|ฯฯ|ๆๆ|[^ฤ]ๅ|ฤ[ะ-ูๆ-์]|[ัี-ืู]์| {2,}|\({2,}|\){2,}|\""{2,}|'{2,}|[\u201C\u201D]{2,}|, *(และ|หรือ)|[ฺํ-๏๚๛๐-๙!?^|—_]|ร้อยละ *\d+ *%|([^\sล]|[^ฯ]ล|^)ฯ\S|(^|\s)[ะ-ู็-์]|\D:[^\s/]|\S:[^\d/])""", re.UNICODE)

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
st.title("Thai Spellchecker for DOCX")
st.write("🔍 Upload a `.docx` file to find and highlight:")
st.markdown("""
- ❌ Thai spelling errors (🔴 red)<br>
- ⚠️ Unexpected Thai dot ◌ฺ (🟠 orange)<br>
- ⚠️ Misused apostrophes `'` (🟣 purple)<br>
- ⚠️ Invalid period use `.` (🔵 blue)<br>
- ⚠️ Common error words (🟡 yellow)<br>
- ⚠️ RegEx error (🟧 bright orange)
""", unsafe_allow_html=True)

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


def find_common_errors(text):
    return [word for word in COMMON_ERRORS if word in text]


def find_regex_errors(text):
    matches = []
    for m in REGEX_ERROR_PATTERN.finditer(text):
        start = m.start()
        value = m.group()
        # Skip if within the first 15 characters or match is only Thai numerals
        if start < 15:
            continue
        if all(c in "๐๑๒๓๔๕๖๗๘๙" for c in value.strip()):
            continue
        matches.append(value)
    return matches

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
        regex_errors = find_regex_errors(text)
        marked = safe_check(text)

        if ("<คำผิด>" in marked or has_phinthu or has_apostrophe or
                invalid_periods or common_errors or regex_errors):
            results.append({
                "line_no": i + 1,
                "original": text,
                "marked": marked,
                "has_phinthu": has_phinthu,
                "has_apostrophe": has_apostrophe,
                "invalid_periods": invalid_periods,
                "common_errors": common_errors,
                "regex_errors": regex_errors
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

        raw_text = raw_text.replace("<คำผิด>", "[[WRONG_OPEN]]").replace("</คำผิด>", "[[WRONG_CLOSE]]")
        safe_text = escape(raw_text)
        safe_text = safe_text.replace("[[WRONG_OPEN]]", "<mark style='background-color:#ffcccc;'>")
        safe_text = safe_text.replace("[[WRONG_CLOSE]]", "</mark>")

        safe_text = safe_text.replace(escape(PHINTHU), mark(PHINTHU, "#ffb84d"))

        safe_text = re.sub(
            r"(>[^<]*)'([^<]*<)",
            lambda m: f"{m.group(1)}<mark style='background-color:#d5b3ff;'>'</mark>{m.group(2)}",
            safe_text
        )

        safe_text = re.sub(
            r"(?<!\w)(\.)(?!\w)",
            lambda m: mark(".", "#add8e6"),
            safe_text
        )

        for word in COMMON_ERRORS:
            safe_text = safe_text.replace(
                escape(word),
                mark(word, "#ffff66")
            )

        for err in item.get("regex_errors", []):
            safe_text = safe_text.replace(
                escape(err),
                mark(err, "#ffa500")
            )

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

        if item.get("regex_errors"):
            html += f"<span style='color:#ff6600;'>⚠️ RegEx error(s) detected</span><br>"

        html += f"<code style='color:gray;'>{original}</code><br>"
        html += f"<div style='margin-top:0.5em;font-size:1.1em;'>{safe_text}</div></div>"

    return html


# Main logic
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
            st.success("✅ No typos, apostrophes, ◌ฺ characters, invalid periods, common errors, or regex issues found!")
