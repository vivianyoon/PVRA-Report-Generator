# PVRA Streamlit App (English + Myanmar Kobo Export)

This version supports both:
- Kobo Excel export with English labels
- Kobo Excel export with Myanmar labels

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## What changed
- Auto-detects workbook language from the main sheet
- Accepts Myanmar-labeled columns and repeat sheets
- Generates Word reports in the same language style as the uploaded file
- Still uses `_index` and `_parent_index` to link village rows with repeat sheets
- Ignores `attachment_repeat`

## Note on Myanmar display
The app writes Unicode Myanmar text into the Word documents. For the best appearance, open the files on a system that has Myanmar-compatible fonts such as **Noto Sans Myanmar** installed.
