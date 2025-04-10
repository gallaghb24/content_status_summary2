import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Content Brief Summary", layout="wide")
st.title("📋 Content Brief Summary Generator")

uploaded_file = st.file_uploader("Upload your Production_Lines.xlsx file", type=["xlsx"])

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

        pivot = status_counts.pivot_table(
            index=['project_ref', 'project_description', 'project_owner', 'event_name'],
            columns='content_brief_status',
            values='line_count',
            fill_value=0
        ).reset_index()

        # Standardise column names again after pivot
        pivot.columns.name = None
        pivot.columns = [str(col).strip().replace(" ", "_").lower() for col in pivot.columns]

        # Define merged status groups
        awaiting_brief_statuses = [
            'draft', 'saved', 'awaiting_agency_briefs'
        ]
        awaiting_amends_statuses = [
            'awaiting_artwork_amends', 'client_rejected_artwork',
            'itg_rejected_artwork', 'rejected_artwork'
        ]

        # Add merged columns only from existing status columns
        existing_brief_statuses = [col for col in awaiting_brief_statuses if col in pivot.columns]
        pivot['awaiting_brief'] = pivot[existing_brief_statuses].sum(axis=1, min_count=1)

        existing_amends_statuses = [col for col in awaiting_amends_statuses if col in pivot.columns]
        pivot['awaiting_artwork_amends'] = pivot[existing_amends_statuses].sum(axis=1, min_count=1)

        # Recalculate number of lines directly from source
        line_counts = df.groupby('project_ref')['brief_ref'].count().to_dict()
        pivot['no_of_lines'] = pivot['project_ref'].map(line_counts)

        # Calculate % completed
        pivot['%_completed'] = ((pivot.get('completed', 0) / pivot['no_of_lines']) * 100).round(0).astype(int).astype(str) + '%'

        # Remove columns where all values are 0 (excluding key columns)
        core_cols = ['project_ref', 'project_description', 'project_owner', 'event_name']
        known_summary_cols = core_cols + ['no_of_lines', 'completed', '%_completed', 'awaiting_brief', 'awaiting_artwork_amends']
        non_zero_cols = pivot.loc[:, (pivot != 0).any(axis=0)].columns.tolist()
        display_cols = [col for col in pivot.columns if col in known_summary_cols or col in non_zero_cols]

        final_summary = pivot[display_cols]

        # Add check column: sum of status columns should equal number of lines
        status_cols = [col for col in final_summary.columns if col not in core_cols + ['no_of_lines', '%_completed'] and final_summary[col].dtype in [int, float]]
        final_summary['check_total'] = final_summary[status_cols].sum(axis=1)
        final_summary['check_passes'] = final_summary['check_total'] == final_summary['no_of_lines']

        st.success("✅ Summary generated!")
        st.dataframe(final_summary, use_container_width=True)

        # Create an XLSX download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_summary.to_excel(writer, index=False, sheet_name='Summary')
        output.seek(0)

        st.download_button("📥 Download Summary as XLSX", output, "summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
