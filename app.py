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
    r"\b[0-9]+\.",              # Arabic numeral lists: 1., 2.
    r"\b[ก-ฮ]\.",               # Thai alphabetical lists: ก., ข.
    r"\b[๐-๙]+\.",              # Thai numeral lists: ๒., ๓.
    r"\b[๐-๙]{1,2}\.[๐-๙]{1,2}",# Thai time: ๑๐.๑๐
    r"\bพ\.ศ\.",                # พ.ศ.
    r"\bค\.ศ\.",                # ค.ศ.
    r"\.{3,}"                   # Ellipses: ..., ..........
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
- ⚠️ Non-consecutive "L" numbers (🟢 green)
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Choose a Word document", type="docx")


# Helper functions
def find_invalid_periods(text):
    invalid_indices = []
    for match in re.finditer(r"\.", text):
        is_valid = False
        for pattern in VALID_PERIOD_PATTERns:
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
        # This check attempts to see if the spell checker drastically changed the text,
        # which might indicate an issue with the spell checker itself or the text.
        # It's a heuristic and might need fine-tuning.
        clean_marked = marked.replace("<คำผิด>", "").replace("</คำผิด>", "")
        if len(clean_marked) < len(text) * 0.5: # If more than 50% was marked, it might be an issue.
             return text # Return original if the spellcheck seems overly aggressive
        return marked
    except Exception:
        # Fallback if thaispellcheck encounters an unhandled error
        return text

def find_l_errors(paragraphs_with_l):
    """
    Detects non-consecutive 'L' numbers in a list of (line_no, l_number) tuples.
    Returns a list of line numbers where an 'L' error is detected.
    """
    errors = []
    expected_l = None

    for line_no, current_l in paragraphs_with_l:
        if expected_l is None:
            expected_l = current_l
        elif current_l != expected_l:
            errors.append(line_no)
        expected_l += 1
    return errors


def check_docx(file):
    doc = docx.Document(file)
    paragraphs = doc.paragraphs
    total = len(paragraphs)
    results = []
    l_paragraphs = [] # To store (line_no, l_number) for L-sequence checking

    progress_bar = st.progress(0, text="Processing...")

    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue

        # Check for L followed by a number at the beginning of the paragraph
        l_match = re.match(r"^[Ll](\d+):", text)
        if l_match:
            l_number = int(l_match.group(1))
            l_paragraphs.append((i + 1, l_number))

        has_phinthu = PHINTHU in text
        has_apostrophe = "'" in text
        invalid_periods = find_invalid_periods(text)
        common_errors = find_common_errors(text)
        marked = safe_check(text)

        # We always want to add a result entry for a line if there's any potential issue.
        # This includes L matches, even if they aren't sequence errors yet,
        # so the 'l_match' info is carried over.
        if "<คำผิด>" in marked or has_phinthu or has_apostrophe or invalid_periods or common_errors or l_match:
            results.append({
                "line_no": i + 1,
                "original": text,
                "marked": marked, # This already has <คำผิด> marks
                "has_phinthu": has_phinthu,
                "has_apostrophe": has_apostrophe,
                "invalid_periods": invalid_periods,
                "common_errors": common_errors,
                "l_match": l_match # Store the match object for potential highlighting
            })

        progress = int((i + 1) / total * 100)
        progress_bar.progress(progress, text=f"Processing paragraph {i + 1} of {total} ({progress}%)")

    progress_bar.empty()

    # After processing all paragraphs, check for L-sequence errors
    l_sequence_errors_lines = find_l_errors(l_paragraphs)
    for result_item in results: # Renamed loop variable to avoid conflict
        if result_item["line_no"] in l_sequence_errors_lines:
            result_item["l_sequence_error"] = True
        else:
            result_item["l_sequence_error"] = False

    return results


def render_html(results):
    def escape(text): return html_lib.escape(text)

    def mark(text, color):
        return f"<mark style='background-color:{color};'>{text}</mark>" # Text is already escaped or a mark itself

    html = "<style> mark { padding: 2px 4px; border-radius: 3px; } </style>"

    for item in results:
        line_no = item["line_no"]
        original = escape(item["original"])
        processed_text = item["marked"] # Start with text that already has <คำผิด> marks

        # Step 1: Handle spellcheck marks (already in processed_text from safe_check)
        # Replace <คำผิด> tags with specific highlight marks
        processed_text = processed_text.replace("<คำผิด>", "<mark style='background-color:#ffcccc;'>")
        processed_text = processed_text.replace("</คำผิด>", "</mark>")

        # Step 2: Escape the entire string, except for the already inserted <mark> tags
        # This requires a more careful approach. We can split and escape.
        parts = re.split(r'(<mark[^>]*>.*?<\/mark>)', processed_text)
        safe_parts = []
        for part in parts:
            if part.startswith('<mark') and part.endswith('</mark>'):
                safe_parts.append(part) # Already a mark, keep as is
            else:
                safe_parts.append(escape(part)) # Escape regular text
        safe_text = "".join(safe_parts)


        # Now, apply other highlights, making sure not to double-mark within existing <mark> tags.
        # The key is to avoid using variable-width lookbehinds.
        # Instead, we can process the text and keep track of marked segments or
        # use a more targeted replacement.

        # Step 3: Highlight ◌ฺ (Phinthu)
        # We need to ensure we're not inside an existing <mark> tag.
        # This is tricky with simple re.sub and no variable-width lookbehind.
        # A workaround is to replace characters that are NOT part of a mark.
        # For simplicity, we'll assume phinthu is unlikely to be part of a spelling error.
        safe_text = safe_text.replace(escape(PHINTHU), mark(escape(PHINTHU), "#ffb84d"))


        # Step 4: Highlight apostrophes
        # Apostrophes are typically standalone, so a simple replace on the escaped text is usually fine.
        # If an apostrophe appears inside a highlighted common error word, it won't be re-marked if we're careful.
        # The most robust way is to find original positions and then insert marks.
        # For now, let's try a regex that doesn't use lookbehinds for the mark itself.
        # We target isolated apostrophes or those within plain text.
        # This is a bit of a compromise given the regex limitations.
        # A more robust approach might involve processing the original text and building the HTML.
        safe_text = re.sub(
            r"((?:[^<]|<(?!\/?mark)[^>]*>)*?)'(?!<)", # Matches ' not immediately followed by <
            r"\1" + mark("'", "#d5b3ff"),
            safe_text
        )


        # Step 5: Highlight invalid periods
        # Similar to apostrophes, target periods that are not already part of a mark.
        # This re.sub will find periods that are not inside <...> tags (which includes <mark> tags).
        # This is still a heuristic, but safer than variable-width lookbehind.
        safe_text = re.sub(
            r"(\.)(?![^<]*>)", # Match a dot not followed by any non-< character then a >
            lambda m: mark(m.group(1), "#add8e6"),
            safe_text
        )


        # Step 6: Highlight common errors
        # This is where the original PatternError occurred.
        # The safest way to highlight common errors without conflicting with existing marks
        # (especially spellcheck marks) is to prioritize. Spellcheck comes first.
        # Then, if a common error word is NOT already surrounded by <mark> tags, mark it.
        # This requires more complex string manipulation or a parsing approach,
        # but for simplicity, we can try to replace only if the word isn't already marked.
        # This is still a challenge with `re.sub` and no variable-width lookbehind.
        # A practical approach for this is to operate on the original text, then substitute.
        # However, since `item["marked"]` already has spellcheck errors, let's keep working on `safe_text`
        # but with a more careful replacement.

        # A simpler (but less perfect) approach for common errors without complex lookbehinds:
        # Iterate and replace. This might mark parts of words already inside other marks if not careful.
        # The best way to avoid this is to ensure the replacement happens only on plain text.
        # Let's try to rebuild `safe_text` in a loop to handle this more safely.

        # Re-evaluating the approach for common errors:
        # Instead of regex for common errors on `safe_text`, let's do this before `escape()` and after `<คำผิด>`.
        # This means modifying `processed_text` before it's fully escaped.

        # Let's revert to a slightly different approach for common errors to avoid the regex issue:
        # Common errors should be highlighted BEFORE general escaping, but AFTER thaispellcheck.
        temp_text_for_common_errors = item["marked"] # This has <คำผิด>
        for error_word in COMMON_ERRORS:
            # Only replace if the word is NOT within <คำผิด> tags.
            # This is still complex with simple regex.
            # A common workaround is to use a replacer function that checks context.
            # However, the easiest is to just let spellcheck take precedence, and only highlight if not a spellcheck error.
            # Given the previous issue, let's prioritize <คำผิด> and then simply replace common errors in the escaped text,
            # acknowledging that if a common error is also a <คำผิด>, the red highlight will take precedence.
            # We already have spellcheck marks. Now we're looking for common errors in the *already escaped and spell-checked* text.
            # If `mark` calls `escape` internally, we need to pass the raw `error_word` and let `mark` handle escaping.

            # We need to find `error_word` not inside an existing `<mark>` tag.
            # This is the tricky part without fixed-width lookbehinds.
            # Let's try a regex that matches `error_word` unless it's preceded by `style='background-color` or followed by `</mark>`.
            # This is still not perfect.

            # Alternative for common errors:
            # We can use a simpler `re.sub` without lookbehinds that would cause fixed-width errors,
            # and rely on the order of operations. Spellcheck has already marked its errors.
            # Now, for common errors, we can simply replace them. If a common error word
            # is *part* of a spell-checked word, the `thaispellcheck` highlight will likely cover it.
            # If it's a standalone common error, it will get highlighted.

            # Simple replacement for common errors (might overlap if not careful):
            # The `mark` function already escapes the text, so we pass the raw `error_word`
            safe_text = safe_text.replace(
                escape(error_word), # Escaped version of the error word
                mark(escape(error_word), "#ffff66") # Mark the escaped version
            )

        # Step 7: Highlight L-sequence errors
        if item.get("l_sequence_error") and item["l_match"]:
            l_full_text = item["l_match"].group(0) # e.g., "L1:"
            # We need to escape `l_full_text` before replacing in `safe_text`
            safe_text = safe_text.replace(
                escape(l_full_text),
                mark(escape(l_full_text), "#90ee90") # Light green for L errors
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

        if item.get("l_sequence_error"):
            html += f"<span style='color:#228b22;'>⚠️ Found non-consecutive 'L' number.</span><br>"


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
            st.success("✅ No typos, apostrophes, ◌ฺ characters, invalid periods, common errors, or 'L' sequence issues found!")
