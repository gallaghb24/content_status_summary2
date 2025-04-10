import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Event Artwork Status Report", layout="wide")
st.title("📊 Event Artwork Status Report")

st.markdown("""
### 📥 How to prepare your file for the Event Artwork Status Report

1. Go to the [**Production Lines report**](https://superdrug.aswmediacentre.com/Reports/Reports/CustomReport?reportId=2) in Media Centre.  
2. Type the **exact name of the Event** you want to report on (e.g. *Event 6 2025*), or choose a **date range** if you'd like to include multiple Events.  
3. Click **Search** to generate the results.  
4. Once the data loads, click the **Excel icon** to export the file. Then click **Export** on the pop-up window.  
5. When the file has downloaded, **open it** and then **Save As** an `.xlsx` Excel file (not `.xls` or `.csv`).  
6. You're now ready — upload the `.xlsx` file using the uploader below.
""")

uploaded_file = st.file_uploader("Choose a file to upload", type=["xlsx"])

if uploaded_file:
    # Load data
    try:
        df = pd.read_excel(uploaded_file, sheet_name='general_report', header=1)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
    else:
        df.columns = [str(col).strip().replace(" ", "_").lower() for col in df.columns]

        # Group and count status lines
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
        ordered_cols = core_cols + [
            'awaiting_brief', 'awaiting_artwork', 'awaiting_artwork_amends',
            'itg_approve_artwork', 'approve_artwork', 'not_applicable',
            'completed', 'no_of_lines', '%_completed'
        ]
        additional_cols = [col for col in pivot.columns if col not in ordered_cols and col not in excluded_from_display and pivot[col].sum() > 0]
        ordered_cols += additional_cols

        final_summary = pivot[[col for col in ordered_cols if col in pivot.columns]].copy()

        st.success("✅ Done!")
        st.dataframe(final_summary, use_container_width=True)

        # Format headers: convert to Proper Case and remove underscores
        formatted_headers = [col.replace("_", " ").title() for col in final_summary.columns]

        # Filename with timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"event_artwork_status_report-{timestamp}.xlsx"

        # Create Excel with formatting and raw data sheet
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_summary.to_excel(writer, index=False, sheet_name='Summary', header=formatted_headers)
            df.to_excel(writer, index=False, sheet_name='Raw Data')

            workbook = writer.book
            worksheet = writer.sheets['Summary']

            # Apply conditional formatting for % Completed
            percent_col_index = formatted_headers.index('% Completed')
            worksheet.conditional_format(1, percent_col_index, len(final_summary), percent_col_index, {
                'type': '3_color_scale',
                'min_color': "#F8696B",  # Red
                'mid_color': "#FFEB84",  # Yellow/Orange
                'max_color': "#63BE7B"   # Green
            })

            # Auto column widths
            for i, col in enumerate(formatted_headers):
                max_len = max(final_summary[col.lower().replace(" ", "_")].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)

        output.seek(0)
        st.download_button("📥 Download Report", output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
