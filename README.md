# Question Bank CSV Converter

Streamlit app to paste questions or upload a `.docx` Word document, optionally upload an Excel template, review extracted questions and download CSV.

## Standard CSV header order

`question type`, `question group`, `question level`, `question`, `marks`, `option_1`, `option_2`, `option_3`, `option_4`, `answer`

Headers use spaces, not underscores. The `question` column comes before `marks`.

## Defaults
- question type: `single_choice`
- question group: `119`
- question level: `easy`
- marks: `1`

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud
Upload `app.py`, `requirements.txt`, and `README.md` to GitHub, then deploy at https://share.streamlit.io/ with `app.py` as the main file.

## Answer mapping
Option A → `option_1`; Option B → `option_2`; Option C → `option_3`; Option D → `option_4`.
