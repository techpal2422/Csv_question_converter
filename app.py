import re
import csv
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="Question Bank CSV Converter", page_icon="📝", layout="wide")
st.title("📝 Question Bank CSV Converter")
st.write("Paste questions or upload a Word document, review the extracted data, and download a CSV.")

# Backend header names use spaces, not underscores. The question column comes before marks.
COLUMNS = [
    "question type",
    "question group",
    "question level",
    "question",
    "marks",
    "option 1",
    "option 2",
    "option 3",
    "option 4",
    "answer",
]

with st.sidebar:
    st.header("Default values")
    q_type = st.text_input("question type", "single_choice")
    q_group = st.text_input("question group", "119")
    q_level = st.text_input("question level", "easy")
    marks = st.text_input("marks", "1")

mode = st.radio("Choose input method", ["Paste text", "Upload Word document"], horizontal=True)
pasted = ""
docx_file = None

if mode == "Paste text":
    pasted = st.text_area(
        "Paste your questions here",
        height=320,
        placeholder="""1. What is the capital of Nigeria?
Option A: Accra
Option B: Lagos
Option C: Abuja
Option D: Kano
Answer: Option C"""
    )
else:
    docx_file = st.file_uploader("Upload Word document (.docx)", type=["docx"])

template = st.file_uploader("Optional Excel template (.xlsx)", type=["xlsx"])

st.markdown(
    "**Expected format:** numbered question, `Option A: ...`, `Option B: ...`, "
    "`Option C: ...`, `Option D: ...`, `Answer: Option A/B/C/D`."
)

def read_docx(file):
    doc = Document(file)
    bits = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            vals = [c.text.strip().replace("\n", " ") for c in row.cells]
            if any(vals):
                bits.append(" ".join(v for v in vals if v))
    return "\n".join(bits)

def parse(raw):
    raw = raw.replace("\r\n", "\n").replace("\r", "\n").replace("\u00a0", " ")
    starts = list(re.finditer(r"(?im)^\s*(?:(?:question)\s*)?\d+\s*[.):\-]\s*", raw))
    if starts:
        chunks = [
            raw[m.start(): starts[i + 1].start() if i + 1 < len(starts) else len(raw)].strip()
            for i, m in enumerate(starts)
        ]
    else:
        chunks = [raw.strip()] if raw.strip() else []

    rows = []
    for chunk in chunks:
        body = re.sub(r"(?is)^\s*(?:(?:question)\s*)?\d+\s*[.):\-]\s*", "", chunk)
        opt_matches = list(re.finditer(
            r"(?im)(?:^|\n)\s*option\s*([A-D])\s*[:.)\-]\s*", body
        ))
        if len(opt_matches) < 4:
            opt_matches = list(re.finditer(
                r"(?i)\boption\s*([A-D])\s*[:.)\-]\s*", body
            ))
        if not opt_matches:
            continue

        ans_match = re.search(
            r"(?i)\banswer\s*[:\-]\s*(?:option\s*)?([A-D])\b", body
        )
        question = body[:opt_matches[0].start()].strip(" \n:-")
        options = {}
        for i, match in enumerate(opt_matches):
            end = opt_matches[i + 1].start() if i + 1 < len(opt_matches) else len(body)
            value = re.split(r"(?i)\banswer\s*[:\-]", body[match.end():end])[0]
            options[match.group(1).upper()] = value.strip(" \n\t;|")

        letter = ans_match.group(1).upper() if ans_match else ""
        answer = f"option {ord(letter) - ord('A') + 1}" if letter else ""
        rows.append({
            "question type": q_type,
            "question group": q_group,
            "question level": q_level,
            "question": question,
            "marks": marks,
            "option 1": options.get("A", ""),
            "option 2": options.get("B", ""),
            "option 3": options.get("C", ""),
            "option 4": options.get("D", ""),
            "answer": answer,
        })
    return rows

if st.button("Extract questions", type="primary", use_container_width=True):
    source = pasted
    if mode == "Upload Word document":
        if docx_file is None:
            st.warning("Please upload a .docx Word document first.")
            source = ""
        else:
            try:
                source = read_docx(docx_file)
            except Exception as exc:
                st.error(f"Could not read the Word document: {exc}")
                source = ""

    if source.strip():
        st.session_state["parsed_rows"] = parse(source)
    else:
        st.session_state.pop("parsed_rows", None)
        if mode == "Paste text":
            st.warning("Please paste your questions first.")

if "parsed_rows" in st.session_state:
    rows = st.session_state["parsed_rows"]
    if not rows:
        st.error("No questions detected. Use numbered questions with Option A–D and Answer labels.")
    else:
        st.subheader("Review and edit")
        df = pd.DataFrame(rows, columns=COLUMNS)
        st.metric("Questions detected", len(df))
        edited = st.data_editor(
            df,
            key="question_editor",
            use_container_width=True,
            num_rows="dynamic",
        )

        output = edited
        if template is not None:
            try:
                template.seek(0)
                template_headers = list(pd.read_excel(template).columns)
                # Normalize template headers for matching, but output standardized spaced headers
                # when the template uses equivalent underscore-style names.
                def norm_header(value):
                    return re.sub(r"\s+", " ", str(value).strip().lower().replace("_", " "))

                source_map = {norm_header(c): c for c in edited.columns}
                output = pd.DataFrame()
                for header in template_headers:
                    key = norm_header(header)
                    source_col = source_map.get(key)
                    output[header] = edited[source_col] if source_col is not None else ""
                st.info("CSV follows the Excel template's exact header spelling and order.")
            except Exception as exc:
                st.warning(f"Could not read Excel template; using standard headers. Details: {exc}")
                output = edited

        csv_data = output.to_csv(index=False, quoting=csv.QUOTE_MINIMAL).encode("utf-8-sig")
        st.download_button(
            "⬇️ Download CSV",
            data=csv_data,
            file_name="converted_questions.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
        )

st.caption(
    "Standard header order: question type, question group, question level, question, marks, "
    "option 1, option 2, option 3, option 4, answer. Headers use spaces rather than underscores. "
    "Answer mapping: Option A → option 1, B → option 2, C → option 3, D → option 4."
)
