import streamlit as st
import functions as f
import os
import datetime

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

# Input for document register download folder
default_doc_path = os.path.expanduser("~/Downloads")
doc_folder = st.text_input(
    label="Select folder to download document register",
    value=default_doc_path,
    help="Enter the folder path where the document register Excel file will be saved."
)

# Validate document register folder
if doc_folder:
    if not os.path.isdir(doc_folder) or not os.access(doc_folder, os.W_OK):
        st.error("Invalid or non-writable folder for document register. Please choose a valid folder.")
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
    with st.form(key='submission_period_form'):
        submission_period = st.date_input(
            label='Select a period to extract submissions from the document register',
            help='Selected dates are inclusive i.e. 00:00 on the start date to 23:59 on the end date',
            value=[
                st.session_state.doc_register['submission_date'].max() - datetime.timedelta(days=3),
                st.session_state.doc_register['submission_date'].max()
            ],
            max_value=st.session_state.doc_register['submission_date'].max(),
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
                # Input for forepage output folder
                output_folder = st.text_input(
                    label="Select folder to save generated forepages",
                    value=default_doc_path,
                    help="Enter the folder path where the generated Word documents will be saved."
                )
                # Validate output folder
                if output_folder and (not os.path.isdir(output_folder) or not os.access(output_folder, os.W_OK)):
                    st.error("Invalid or non-writable folder for forepages. Please choose a valid folder.")
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