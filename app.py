import pandas as pd
import streamlit as st

from report_generator import (
    APP_TITLE,
    MAIN_SHEET_CANDIDATES,
    REPEAT_SHEETS,
    build_zip,
    detect_language,
    find_main_sheet,
    generate_reports,
    getcol,
    load_workbook,
)

st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption(
    "Upload Kobo Excel export. The app auto-detects English or Myanmar column labels, links repeat sheets using _index and _parent_index, and generates one Word report per village."
)

with st.expander("Expected workbook structure", expanded=False):
    st.markdown(
        """
        - Main village sheet: **PVRA Finding** or **PVRA Village Assessment Form**
        - Repeat sheets linked by **_parent_index**:
          - livelihood_repeat
          - agriculture_repeat
          - livestock_repeat
          - hazard_repeat
          - seasonal_calendar_repeat
          - priority_ranking_repeat
        - attachment_repeat is ignored.
        - Supports both **old structure** and **new structure**, in **English** or **Myanmar** export from Kobo.
        - In the new structure, some hazard summary fields are on the main sheet and hazard events stay in **hazard_repeat**.
        """
    )

uploaded_file = st.file_uploader("Upload Kobo Excel export (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        workbook = load_workbook(uploaded_file)
        main_sheet = find_main_sheet(workbook)
        if not main_sheet:
            st.error(f"Main sheet is missing. Expected one of: {', '.join(MAIN_SHEET_CANDIDATES)}")
            st.stop()

        main_df = workbook[main_sheet].copy()
        lang = detect_language(main_df)
        village_count = len(main_df)
        st.success(
            f"Workbook loaded successfully. Main sheet: {main_sheet}. Found {village_count} village submission(s). Detected language: {'Myanmar' if lang == 'my' else 'English'}."
        )

        preview_cols = [
            c
            for c in [
                "_index",
                getcol(main_df, "village_name"),
                getcol(main_df, "township_name"),
                getcol(main_df, "district_name"),
                getcol(main_df, "assessment_date"),
            ]
            if c
        ]
        st.subheader("Village preview")
        st.dataframe(main_df[preview_cols].copy(), use_container_width=True)

        stats = []
        for sheet_name, label_name in REPEAT_SHEETS.items():
            df = workbook.get(sheet_name, pd.DataFrame())
            stats.append({"Sheet": sheet_name, "Section": label_name, "Rows": len(df)})
        st.subheader("Repeat sheet summary")
        st.dataframe(pd.DataFrame(stats), use_container_width=True)

        if st.button("Generate Word reports", type="primary"):
            reports = generate_reports(workbook)
            zip_bytes = build_zip(reports)
            st.success(f"Generated {len(reports)} Word report(s).")
            st.info(
                "Myanmar text is supported in the generated Word files. For best display, open the report on a device with Myanmar-compatible fonts such as Noto Sans Myanmar installed."
            )
            st.download_button(
                "Download all village reports (.zip)",
                data=zip_bytes,
                file_name="pvra_village_reports.zip",
                mime="application/zip",
            )

            st.subheader("Individual downloads")
            cols = st.columns(2)
            for i, item in enumerate(reports):
                with cols[i % 2]:
                    st.download_button(
                        label=f"Download {item['village_name']}.docx",
                        data=item["bytes"],
                        file_name=item["village_name"].replace("/", "_") + ".docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{i}",
                    )
    except Exception as exc:
        st.exception(exc)
