import streamlit as st
import functions as f
import os
import datetime
import logging

# Set up logging for production debugging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def downloadOnwards(doc_register_path):
    try:
        f.downloadDocRegister(output_path=doc_register_path)
        st.session_state.doc_register = f.readDocRegister(path=doc_register_path)
    except Exception as e:
        st.error(f"Error downloading or reading document register: {str(e)}")

if 'doc_register' not in st.session_state:
    st.session_state.doc_register = None

if 'checkboxes' not in st.session_state:
    st.session_state.checkboxes = None

st.title('Weekly Report Forepage Autofill')
st.subheader('Lantau Portfolio Project')

# Use a production-safe default path
default_doc_path = os.path.abspath(os.getcwd())  # Current working directory
if not os.access(default_doc_path, os.W_OK):
    default_doc_path = "/tmp"  # Fallback to /tmp if cwd is not writable
logger.info(f"Default document register path: {default_doc_path}")

doc_folder = st.text_input(
    label="Select folder to download document register",
    value=default_doc_path,
    help="Enter a writable folder path where the document register Excel file will be saved."
)

# Validate and create document register folder
if doc_folder:
    try:
        os.makedirs(doc_folder, exist_ok=True)  # Create folder if it doesn't exist
        if not os.path.isdir(doc_folder):
            st.error("Invalid folder for document register. Please choose a valid directory.")
            doc_folder = None
        elif not os.access(doc_folder, os.W_OK):
            st.error("Selected folder is not writable. Please choose a folder with write permissions.")
            doc_folder = None
        else:
            logger.info(f"Validated document register folder: {doc_folder}")
    except Exception as e:
        st.error(f"Error accessing folder: {str(e)}")
        logger.error(f"Folder access error: {str(e)}")
        doc_folder = None
else:
    st.error("Please specify a folder for the document register.")
    doc_folder = None

# Generate a unique filename for the document register
doc_register_filename = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "_nex1110-doc-register.xlsx"
doc_register_path = os.path.join(doc_folder, doc_register_filename) if doc_folder else None

scan_doc_register = st.button(
    label='Scan document register',
    disabled=not doc_folder,
    on_click=downloadOnwards,
    args=(doc_register_path,)
)

submission_period = None
if st.session_state.doc_register is not None:
    # Check if doc_register is empty or has invalid dates
    if st.session_state.doc_register.empty or st.session_state.doc_register['submission_date'].isna().all():
        st.error("Document register is empty or contains no valid submission dates. Please check the data source.")
    else:
        with st.form(key='submission_period_form'):
            # Calculate max date (min_value removed per your fix)
            max_date = st.session_state.doc_register['submission_date'].max()
            default_start = max_date - datetime.timedelta(days=3)
            # Convert to date objects if they are datetime
            if isinstance(default_start, datetime.datetime):
                default_start = default_start.date()
            if isinstance(max_date, datetime.datetime):
                max_date = max_date.date()

            submission_period = st.date_input(
                label='Select a period to extract submissions from the document register',
                help='Selected dates are inclusive i.e. 00:00 on the start date to 23:59 on the end date',
                value=[default_start, max_date],
                max_value=max_date,
                format='DD/MM/YYYY'
            )
            confirm_period = st.form_submit_button(
                label='Confirm period'
            )

    if confirm_period:
        doc_register_for_period = st.session_state.doc_register[
            (st.session_state.doc_register['submission_date'] >= datetime.datetime.combine(submission_period[0], datetime.datetime.min.time())) &
            (st.session_state.doc_register['submission_date'] <= datetime.datetime.combine(submission_period[1], datetime.datetime.min.time()))
        ]
        if doc_register_for_period.empty:
            st.warning(
                body='No submissions found in selected period. Please re-select the period.',
                icon=':material/warning:'
            )
        else:
            with st.form(key='doc_selection_form'):
                st.caption('Select submissions to generate forepages for')
                # Input for forepage output folder (use same default as doc_folder)
                output_folder = st.text_input(
                    label="Select folder to save generated forepages",
                    value=default_doc_path,
                    help="Enter a writable folder path where the generated Word documents will be saved."
                )
                # Validate output folder
                if output_folder:
                    try:
                        os.makedirs(output_folder, exist_ok=True)
                        if not os.path.isdir(output_folder):
                            st.error("Invalid folder for forepages. Please choose a valid directory.")
                            output_folder = None
                        elif not os.access(output_folder, os.W_OK):
                            st.error("Selected folder is not writable. Please choose a folder with write permissions.")
                            output_folder = None
                        else:
                            logger.info(f"Validated forepage output folder: {output_folder}")
                    except Exception as e:
                        st.error(f"Error accessing forepage folder: {str(e)}")
                        logger.error(f"Forepage folder access error: {str(e)}")
                        output_folder = None
                else:
                    st.error("Please specify a folder for the forepages.")
                    output_folder = None

                st.session_state.doc_register = doc_register_for_period
                st.session_state.checkboxes = [
                    st.checkbox(
                        label=doc['report_title'],
                        value=True,
                        help=doc['doc_number']
                    ) for _, doc in st.session_state.doc_register.iterrows()
                ]
                generate_forepages = st.form_submit_button(
                    label='Generate forepages',
                    disabled=not output_folder
                )

                if generate_forepages and output_folder:
                    selected_submissions = st.session_state.doc_register[st.session_state.checkboxes]
                    if selected_submissions.empty:
                        st.warning("No submissions selected. Please select at least one submission.")
                    else:
                        try:
                            f.generateForepages(selected_submissions, output_folder)
                            st.success(f"Forepages generated successfully in {output_folder}!")
                        except Exception as e:
                            st.error(f"Error generating forepages: {str(e)}")
                        st.session_state.doc_register = None
                        st.session_state.checkboxes = None