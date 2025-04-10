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

        full_pivot = status_counts.pivot_table(
            index=['project_ref', 'project_description', 'project_owner', 'event_name'],
            columns='content_brief_status',
            values='line_count',
            fill_value=0
        ).reset_index()

        # Standardise column names again after pivot
        full_pivot.columns.name = None
        full_pivot.columns = [str(col).strip().replace(" ", "_").lower() for col in full_pivot.columns]

        # Work on a copy for display summary
        pivot = full_pivot.copy()

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

        # Track columns used for merging to exclude from display but not from check
        excluded_from_display = set(existing_brief_statuses + existing_amends_statuses)

        # Build display columns
        core_cols = ['project_ref', 'project_description', 'project_owner', 'event_name']
        ordered_cols = core_cols + [
            'awaiting_brief', 'awaiting_artwork', 'awaiting_artwork_amends',
            'itg_approve_artwork', 'approve_artwork', 'not_applicable',
            'completed', 'no_of_lines', '%_completed'
        ]
        additional_cols = [col for col in pivot.columns if col not in ordered_cols and col not in excluded_from_display and pivot[col].sum() > 0]
        ordered_cols += additional_cols

        final_summary = pivot[[col for col in ordered_cols if col in pivot.columns]].copy()

        # Add check column using ALL numeric status columns from the full unfiltered pivot
        status_cols = [col for col in full_pivot.columns if col not in core_cols and full_pivot[col].dtype in [int, float]]
        check_totals = full_pivot[status_cols].sum(axis=1)
        final_summary['check_total'] = check_totals.values
        final_summary['check_passes'] = final_summary['check_total'] == final_summary['no_of_lines']

        st.success("✅ Summary generated!")
        st.dataframe(final_summary, use_container_width=True)

        # Create an XLSX download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_summary.to_excel(writer, index=False, sheet_name='Summary')
        output.seek(0)

        st.download_button("📥 Download Summary as XLSX", output, "summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
