
import io
import zipfile
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from parse_calendar import extract_events

st.set_page_config(page_title="Calendar Parser", layout="wide")
st.title("Calendar PDF Parser")
st.write("Upload a ZIP archive containing calendar PDFs to generate a consolidated, date-sorted table of events.")

uploaded_zip = st.file_uploader("Upload ZIP", type=["zip"], help="ZIP archive containing one or more calendar PDFs")

if uploaded_zip is not None:
    with st.spinner("Processing PDFs..."):
        rows = []
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            with zipfile.ZipFile(io.BytesIO(uploaded_zip.read())) as zf:
                pdf_members = [member for member in zf.namelist() if member.lower().endswith('.pdf')]
                if not pdf_members:
                    st.warning("No PDF files found in the uploaded ZIP archive.")
                else:
                    for member in sorted(pdf_members):
                        if member.endswith('/'):
                            continue
                        target_path = tmpdir_path / Path(member).name
                        with zf.open(member) as source, target_path.open('wb') as target:
                            target.write(source.read())

                        for date_str, event_name, meeting_place, person in extract_events(target_path):
                            rows.append({
                                "date": date_str,
                                "Event Name": event_name,
                                "Meeting Place": meeting_place,
                                "Person": person,
                                "Source PDF": Path(member).name,
                            })

        if rows:
            df = pd.DataFrame(rows)
            df['date'] = pd.to_datetime(df['date'], errors='coerce')
            df = df.sort_values('date').reset_index(drop=True)
            df['date'] = df['date'].dt.date

            st.success(f"Parsed {len(rows)} events from {df['Source PDF'].nunique()} PDF(s).")
            st.dataframe(df, use_container_width=True)

            csv_buffer = io.StringIO()
            df.to_csv(csv_buffer, index=False)
            st.download_button(
                label="Download CSV",
                data=csv_buffer.getvalue(),
                file_name="calendar_events.csv",
                mime="text/csv",
            )
        else:
            st.info("No events were parsed from the uploaded files.")
