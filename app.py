# app.py
import streamlit as st
import docx
import thaispellcheck
from pythainlp.spell import correct
from pythainlp.tokenize import word_tokenize
import html as html_lib
import re

# --- Constants and Configuration ---
PHINTHU = "\u0E3A"
VALID_PERIOD_PATTERNS = [
    r"\b[0-9]+\.",          # Arabic numbered list: 1., 2., etc.
    r"\b[๐-๙]+\.",          # Thai numbered list: ๑., ๒., etc.
    r"\b[ก-ฮ]\.",           # Thai alphabetical list: ก., ข., etc.
    r"\bพ\.ศ\.",             # Buddhist Era
    r"\bค\.ศ\.",             # Christian Era
    r"[๐-๙]{1,2}\.[๐-๙]{1,2}"  # Thai time or decimal number: ๑๒.๓๕
]

# --- UI Setup ---
st.set_page_config(page_title="Thai DOCX Spellchecker", layout="wide")
st.title("Thai Spellchecker for DOCX (Data Wow)")
st.write("🔍 Upload one or more `.docx` files to find and highlight issues.")
st.markdown("""
- 🔥 **High Confidence Error** (🔴 red): Flagged by **both** spellcheck libraries.
- ⚠️ **Low Confidence Error** (🟤 brown): Flagged by **only one** library.
- ⚠️ Unexpected Thai dot `◌ฺ` (🟠 orange).
- ⚠️ Misused apostrophes `'` (🟣 purple).
- ⚠️ Invalid period use `.` (🔵 blue).
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Choose Word documents",
    type="docx",
    accept_multiple_files=True
)


# --- Backend Functions ---

def find_invalid_periods(text):
    """Finds period characters that do not match a list of valid patterns."""
    invalid_indices = []
    for match in re.finditer(r"\.", text):
        is_valid = False
        # Check a small context window around the period
        context_start = max(0, match.start() - 10)
        context_end = min(len(text), match.end() + 10)
        context = text[context_start:context_end]

        for pattern in VALID_PERIOD_PATTERNS:
            # Search for valid patterns within the context window
            for found_pattern in re.finditer(pattern, context):
                # Ensure the period we found is part of the valid pattern
                if match.start() >= context_start + found_pattern.start() and \
                   match.end() <= context_start + found_pattern.end():
                    is_valid = True
                    break
            if is_valid:
                break
        if not is_valid:
            invalid_indices.append(match.start())
    return invalid_indices

def check_docx(file, progress_callback):
    """
    Checks a DOCX file using an optimized "batch" approach and reports
    progress back to the UI via a callback function.
    """
    # --- Step 1: Reading and Initial Tokenization ---
    progress_callback(0, "Step 1/4: Reading document and finding unique words...")
    doc = docx.Document(file)
    if not doc.paragraphs:
        return []
    
    all_para_texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    if not all_para_texts:
        return []
    full_text = " ".join(all_para_texts)
    unique_words = set(word_tokenize(full_text, engine="newmm"))

    # --- Step 2: Checking with thaispellcheck ---
    progress_callback(25, "Step 2/4: Analyzing with the first spellchecker (thaispellcheck)...")
    thaispell_marked_full = thaispellcheck.check(full_text, autocorrect=False)
    doc_misspelled_thaispell = set(re.findall(r"<คำผิด>(.*?)</คำผิด>", thaispell_marked_full))

    # --- Step 3: Checking with pythainlp ---
    progress_callback(50, "Step 3/4: Analyzing with the second spellchecker (pythainlp)...")
    doc_misspelled_pythainlp = set()
    for word in unique_words:
        # A word is misspelled if it's not a number and the library suggests a correction
        if word.strip() and not word.isnumeric() and word != correct(word):
            doc_misspelled_pythainlp.add(word)

    # Determine confidence levels for the entire document
    doc_high_confidence = doc_misspelled_thaispell.intersection(doc_misspelled_pythainlp)
    doc_low_confidence = doc_misspelled_thaispell.symmetric_difference(doc_misspelled_pythainlp)

    # --- Step 4: Compiling Results ---
    progress_callback(90, "Step 4/4: Finalizing and compiling results...")
    results = []
    for i, text in enumerate(all_para_texts):
        # Find which of the pre-identified errors are in this specific paragraph
        para_words = set(word_tokenize(text, engine="newmm"))
        high_confidence_errors = para_words.intersection(doc_high_confidence)
        low_confidence_errors = para_words.intersection(doc_low_confidence)

        # Other stylistic checks
        has_phinthu = PHINTHU in text
        has_apostrophe = "'" in text
        invalid_periods = find_invalid_periods(text)

        # Aggregate results if any issue is found in this paragraph
        if high_confidence_errors or low_confidence_errors or has_phinthu or has_apostrophe or invalid_periods:
            results.append({
                "line_no": i + 1, "original": text,
                "high_confidence_errors": high_confidence_errors, "low_confidence_errors": low_confidence_errors,
                "has_phinthu": has_phinthu, "has_apostrophe": has_apostrophe, "invalid_periods": invalid_periods
            })
            
    progress_callback(100, "Analysis complete!")
    return results

def render_html(results):
    """Renders the list of issues into an HTML string for display."""
    html = "<style> mark { padding: 2px 4px; border-radius: 3px; } </style>"
    for item in results:
        original_escaped = html_lib.escape(item["original"])
        marked_text = original_escaped

        # --- Apply Highlights in Order ---
        # 1. High Confidence Errors (Red)
        for word in item['high_confidence_errors']:
            pattern = r"\b(" + re.escape(html_lib.escape(word)) + r")\b"
            marked_text = re.sub(pattern, r"<mark style='background-color:#ffcccc;'>\1</mark>", marked_text)

        # 2. Low Confidence Errors (Brown/Pink)
        for word in item['low_confidence_errors']:
            pattern = r"\b(" + re.escape(html_lib.escape(word)) + r")\b"
            marked_text = re.sub(pattern, r"<mark style='background-color:#f5cba7;'>\1</mark>", marked_text)

        # 3. Phinthu (Orange)
        marked_text = marked_text.replace(PHINTHU, f"<mark style='background-color:#ffb84d;'>{PHINTHU}</mark>")

        # 4. Apostrophes (Purple)
        marked_text = re.sub(r"'", r"<mark style='background-color:#d5b3ff;'>'</mark>", marked_text)
        
        # 5. Invalid Periods (Blue)
        offset = 0
        for idx in item['invalid_periods']:
            real_idx = idx + offset
            if real_idx > 0 and marked_text[real_idx-1] in ('"', "'"): continue
            
            marked_text = marked_text[:real_idx] + "<mark style='background-color:#add8e6;'>.</mark>" + marked_text[real_idx+1:]
            offset += len("<mark style='background-color:#add8e6;'>.</mark>") - 1

        # --- Build HTML Output for the Item ---
        html += f"<div style='padding:10px;margin-bottom:15px;border:1px solid #ddd;border-radius:5px;'>"
        html += f"<b>Line {item['line_no']}</b><br>"

        if item['high_confidence_errors']:
            html += f"<span style='color:#d00;'>🔥 High Confidence Error(s) found.</span><br>"
        if item['low_confidence_errors']:
            html += f"<span style='color:#804000;'>⚠️ Low Confidence Error(s) found.</span><br>"
        if item['has_phinthu']:
            html += f"<span style='color:#d95f00;'>⚠️ Found unexpected dot (◌ฺ).</span><br>"
        if item['has_apostrophe']:
            html += f"<span style='color:#800080;'>⚠️ Found apostrophe (`'`).</span><br>"
        if item['invalid_periods']:
            html += f"<span style='color:#0055aa;'>⚠️ Found suspicious period (`.`).</span><br>"

        html += f"<hr style='margin: 8px 0; border-top: 1px solid #eee;'>"
        html += f"<code style='color:gray;display:block;margin-bottom:8px;'>{original_escaped}</code>"
        html += f"<div style='font-size:1.1em;'>{marked_text}</div></div>"
    return html

# --- Main Application Logic ---
if uploaded_files:
    for uploaded_file in uploaded_files:
        st.subheader(f"Results for: `{uploaded_file.name}`")

        # Create UI elements for progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()

        # Define the callback function that updates the UI
        def update_progress(percent, message):
            status_text.text(message)
            progress_bar.progress(percent)

        # Run the analysis and pass the callback function
        results = check_docx(uploaded_file, update_progress)

        # Clear the progress elements once done
        status_text.empty()
        progress_bar.empty()

        if results:
            st.markdown(render_html(results), unsafe_allow_html=True)
        else:
            st.success(f"✅ No issues found in {uploaded_file.name}!")
        st.markdown("---")
