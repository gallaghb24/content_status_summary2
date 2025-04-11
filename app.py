import streamlit as st
st.set_page_config(page_title="Event Artwork Status Report", layout="wide")

st.markdown(
    """
    <style>
    h1, h2, h3, h4, h5, h6 {
        visibility: visible;
    }
    .stMarkdown .css-1aumxhk a {
        display: none;
    }
    </style>
    """,
    unsafe_allow_html=True
)

import pandas as pd
from io import BytesIO
from datetime import datetime

st.title("📊 Event Artwork Status Report")

st.markdown("""
### 📥 Instructions

1. Go to the [**Production Lines report**](https://superdrug.aswmediacentre.com/Reports/Reports/CustomReport?reportId=2) in Media Centre.  
2. Type the **exact name of the Event** you want to report on (e.g. *Event 6 2025*), or choose a **date range** if you'd like to include multiple Events.  
3. Click **Search** to generate the results.  
4. Once the data loads, click the **Excel icon** to export the file. On the pop-up, leave it as HTML and click **Export**.  
5. Open the downloaded `.xls` file, re-save it as `.xlsx` Excel file (not `.xls` or `.csv`).  
6. You're now ready — upload the `.xlsx` file using the uploader below.
""")

uploaded_file = st.file_uploader("Choose a file to upload", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name='general_report', header=1)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
    else:
        df.columns = [str(col).strip().replace(" ", "_").lower() for col in df.columns]

        status_counts = (
            df.groupby(['project_ref', 'project_description', 'project_owner', 'event_name', 'content_brief_status'])
              .agg(line_count=('brief_ref', 'count'))
              .reset_index()
        )

        full_pivot = status_counts.pivot_table(
            index=['project_ref', 'project_description', 'project_owner', 'event_name'],
            columns='content_brief_status',
            values='line_count',
            fill_value=0
        ).reset_index()

        full_pivot.columns.name = None
        full_pivot.columns = [str(col).strip().replace(" ", "_").lower() for col in full_pivot.columns]
        pivot = full_pivot.copy()

        awaiting_brief_statuses = ['draft', 'saved', 'awaiting_agency_briefs']
        awaiting_amends_statuses = ['awaiting_artwork_amends', 'client_rejected_artwork', 'itg_rejected_artwork', 'rejected_artwork']

        existing_brief_statuses = [col for col in awaiting_brief_statuses if col in pivot.columns]
        pivot['awaiting_brief'] = pivot[existing_brief_statuses].sum(axis=1, min_count=1)

        existing_amends_statuses = [col for col in awaiting_amends_statuses if col in pivot.columns]
        pivot['awaiting_artwork_amends'] = pivot[existing_amends_statuses].sum(axis=1, min_count=1)

        line_counts = df.groupby('project_ref')['brief_ref'].count().to_dict()
        pivot['no_of_lines'] = pivot['project_ref'].map(line_counts)

        pivot['%_completed'] = ((pivot.get('completed', 0) / pivot['no_of_lines']) * 100).round(0).astype(int).astype(str) + '%'

        excluded_from_display = set(existing_brief_statuses + existing_amends_statuses)

        core_cols = ['project_ref', 'project_description', 'project_owner', 'event_name']

        defined_status_order = [
            'awaiting_brief',
            'awaiting_artwork',
            'awaiting_artwork_amends',
            'itg_approve_artwork',
            'approve_artwork',
            'awaiting_artwork_submission',
            'awaiting_production_ready',
            'itg_rejected_briefs',
            'itg_agency_modifications',
            'agency_modifications',
            'not_applicable',
            'completed'
        ]

        status_cols_in_data = [col for col in defined_status_order if col in pivot.columns]
        extra_status_cols = [
            col for col in pivot.columns
            if col not in core_cols + status_cols_in_data + list(excluded_from_display) + ['no_of_lines', '%_completed']
        ]

        ordered_cols = core_cols + status_cols_in_data + ['no_of_lines', '%_completed'] + extra_status_cols
        final_summary = pivot[[col for col in ordered_cols if col in pivot.columns]].copy()

        formatted_headers = [
            "ITG Approve Artwork" if col == "itg_approve_artwork"
            else "Total Lines" if col == "no_of_lines"
            else col.replace("_", " ").title()
            for col in final_summary.columns
        ]

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"event_artwork_status_report-{timestamp}.xlsx"

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_summary.to_excel(writer, index=False, sheet_name='Summary', header=formatted_headers)

            # Add overall % formula below the table
            total_rows = len(final_summary) + 1  # account for header
            col_map = {h: i for i, h in enumerate(formatted_headers)}
            if 'Total Lines' in col_map and '% Completed' in col_map:
                total_col = col_map['Total Lines']
                percent_col = col_map['% Completed']
                worksheet.write(total_rows + 1, percent_col - 1, 'Overall % Completed')
                formula = f"=ROUND(SUMPRODUCT({chr(65+percent_col)}2:{chr(65+percent_col)}{total_rows},{chr(65+total_col)}2:{chr(65+total_col)}{total_rows})/SUM({chr(65+total_col)}2:{chr(65+total_col)}{total_rows}), 0) & \"%\""
                worksheet.write_formula(total_rows + 1, percent_col, formula, workbook.add_format({"bold": True, "bg_color": "#F0F0F0"}))
                worksheet.write(total_rows + 1, percent_col - 1, 'Overall % Completed', workbook.add_format({"bold": True, "bg_color": "#F0F0F0"}))
            df.to_excel(writer, index=False, sheet_name='Raw Data')

            workbook = writer.book
            worksheet = writer.sheets['Summary']

            if '% Completed' in formatted_headers:
                percent_col_index = formatted_headers.index('% Completed')
                worksheet.conditional_format(1, percent_col_index, len(final_summary), percent_col_index, {
                    'type': '3_color_scale',
                    'min_color': "#F8696B",
                    'mid_color': "#FFEB84",
                    'max_color': "#63BE7B"
                })

            for i, col in enumerate(formatted_headers):
                raw_col_key = col.replace(" ", "_").lower()
                if raw_col_key in final_summary.columns:
                    max_len = max(final_summary[raw_col_key].astype(str).map(len).max(), len(col)) + 2
                else:
                    max_len = len(col) + 2
                worksheet.set_column(i, i, max_len)

        output.seek(0)

        st.markdown("<div style='margin-top: 1.5em; text-align: center;'>", unsafe_allow_html=True)
        st.download_button(
            "✅ Done! Download Report",
            output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False
        )
        st.markdown("</div>", unsafe_allow_html=True)
