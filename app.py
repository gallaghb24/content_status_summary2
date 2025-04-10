import streamlit as st
import pandas as pd

st.set_page_config(page_title="Content Status Summary", layout="wide")
st.title("📊 Content Status Summary")

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

        # Calculate merged columns
        pivot['awaiting_brief'] = (
            pivot.get('draft', 0) +
            pivot.get('saved', 0) +
            pivot.get('awaiting_agency_briefs', 0)
        )

        pivot['awaiting_artwork_amends'] = (
            pivot.get('awaiting_artwork_amends', 0) +
            pivot.get('client_rejected_artwork', 0)
        )

        pivot['no_of_lines'] = pivot.drop(
            columns=['project_ref', 'project_description', 'project_owner', 'event_name', 'completed'], errors='ignore'
        ).sum(axis=1) + pivot.get('completed', 0)

        pivot['%_completed'] = (
            (pivot.get('completed', 0) / pivot['no_of_lines']) * 100
        ).round(0).astype(int).astype(str) + '%'

        # Final column order
        columns_to_display = [
            'project_ref', 'project_description', 'project_owner', 'event_name', 'no_of_lines',
            'awaiting_brief', 'awaiting_artwork', 'awaiting_artwork_amends',
            'itg_approve_artwork', 'approve_artwork', 'not_applicable',
            'completed', '%_completed'
        ]

        columns_to_display = [col for col in columns_to_display if col in pivot.columns]
        final_summary = pivot[columns_to_display]

        st.success("✅ Summary generated!")
        st.dataframe(final_summary, use_container_width=True)

        # Option to download the result
        csv = final_summary.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Summary as CSV", csv, "summary.csv", "text/csv")
