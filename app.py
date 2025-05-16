import streamlit as st
import functions as f
import os
import datetime
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def downloadOnwards(doc_register_path):
    try:
        f.downloadDocRegister(output_path=doc_register_path)
        st.session_state.doc_register = f.readDocRegister(path=doc_register_path)
        logger.info("Document register downloaded and read successfully")
    except Exception as e:
        st.error(f"Error downloading or reading document register: {str(e)}")
        logger.error(f"Document register error: {str(e)}")

# Initialize session state
if 'doc_register' not in st.session_state:
    st.session_state.doc_register = None
if 'stage' not in st.session_state:
    st.session_state.stage = "select_folder"
if 'submission_period' not in st.session_state:
    st.session_state.submission_period = None

st.title('Weekly Report Forepage Autofill')
st.subheader('Lantau Portfolio Project')

# Stage 1: Select Folder
if st.session_state.stage == "select_folder":
    default_doc_path = "/tmp"
    logger.info(f"Default document register path: {default_doc_path}")
    doc_folder = st.text_input(
        label="Select folder to download document register",
        value=default_doc_path,
        help="Enter a writable folder path (e.g., /tmp in production)."
    )
    if doc_folder:
        try:
            os.makedirs(doc_folder, exist_ok=True)
            if not os.path.isdir(doc_folder):
                st.error("Invalid folder. Please choose a valid directory.")
            elif not os.access(doc_folder, os.W_OK):
                st.error("Folder not writable. Please choose a writable folder.")
            else:
                st.session_state.doc_folder = doc_folder
                st.session_state.stage = "scan_register"
        except Exception as e:
            st.error(f"Error accessing folder: {str(e)}")
            logger.error(f"Folder access error: {str(e)}")

# Stage 2: Scan Register
if st.session_state.stage == "scan_register":
    doc_register_filename = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "_nex1110-doc-register.xlsx"
    doc_register_path = os.path.join(st.session_state.doc_folder, doc_register_filename)
    if st.button(label='Scan document register'):
        downloadOnwards(doc_register_path)
        if st.session_state.doc_register is not None:
            st.session_state.stage = "select_period"

# Stage 3: Select Period
if st.session_state.stage == "select_period" and st.session_state.doc_register is not None:
    if st.session_state.doc_register.empty or st.session_state.doc_register['submission_date'].isna().all():
        st.error("Document register is empty or has no valid dates. Check the data source.")
    else:
        with st.form(key='submission_period_form'):
            max_date = st.session_state.doc_register['submission_date'].max()
            default_start = max_date - datetime.timedelta(days=3)
            if isinstance(default_start, datetime.datetime):
                default_start = default_start.date()
            if isinstance(max_date, datetime.datetime):
                max_date = max_date.date()
            submission_period = st.date_input(
                label='Select a period for submissions',
                help='Dates are inclusive (00:00 start to 23:59 end).',
                value=[default_start, max_date],
                max_value=max_date,
                format='DD/MM/YYYY'
            )
            confirm_period = st.form_submit_button(label='Confirm period')
        if confirm_period:
            st.session_state.submission_period = submission_period
            st.session_state.stage = "select_submissions"

# Stage 4: Select Submissions and Generate
if st.session_state.stage == "select_submissions" and st.session_state.doc_register is not None:
    doc_register_for_period = st.session_state.doc_register[
        (st.session_state.doc_register['submission_date'] >= datetime.datetime.combine(st.session_state.submission_period[0], datetime.datetime.min.time())) &
        (st.session_state.doc_register['submission_date'] <= datetime.datetime.combine(st.session_state.submission_period[1], datetime.datetime.min.time()))
    ]
    if doc_register_for_period.empty:
        st.warning('No submissions found. Please re-select the period.')
        st.session_state.stage = "select_period"
    else:
        st.caption('Select submissions to generate forepages for')
        checkboxes = [
            st.checkbox(
                label=doc['report_title'],
                value=True,
                key=f"checkbox_{i}",
                help=doc['doc_number']
            ) for i, (_, doc) in enumerate(doc_register_for_period.iterrows())
        ]
        if st.button(label='Generate forepages'):  # Moved outside form
            selected_submissions = doc_register_for_period.iloc[
                [i for i, checked in enumerate(checkboxes) if checked]
            ]
            if selected_submissions.empty:
                st.warning("No submissions selected. Select at least one.")
            else:
                try:
                    logger.info(f"Generating forepages for {len(selected_submissions)} submissions")
                    zip_data, zip_filename = f.generateForepages(selected_submissions, st.session_state.doc_folder)
                    st.download_button(
                        label="Download Forepages (ZIP)",
                        data=zip_data,
                        file_name=zip_filename,
                        mime="application/zip",
                        help="Download a ZIP file with all forepages."
                    )
                    st.success("Forepages generated! Click to download.")
                except Exception as e:
                    st.error(f"Error generating forepages: {str(e)}")
                    logger.error(f"Forepage generation error: {str(e)}")