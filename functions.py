import streamlit as st
import gdown
import pandas as pd
import datetime
import os
import zipfile
import logging
from docxtpl import DocxTemplate

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def downloadDocRegister(output_path: str):
    url = 'https://drive.google.com/uc?id=1Uvt9HNpA0THgVv42bu8ye-7Zif6VttwyJJ24ywZZk80'
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    gdown.download(url, output_path, quiet=False)

def readDocRegister(path: str) -> pd.DataFrame:
    try:
        doc_register = pd.read_excel(
            io=path,
            header=None,
            names=[
                'report_title',
                'doc_number_pdf',
                'version',
                'submittor',
                'submission_date'
            ],
            usecols='B,I,K,L,M'
        )
        doc_register = doc_register[doc_register['submission_date'].apply(lambda x: isinstance(x, datetime.datetime))]
        doc_register['doc_number'] = doc_register['doc_number_pdf'].apply(lambda x: x[:-4])
        logger.info(f"Read document register with {len(doc_register)} valid entries")
        return doc_register
    except Exception as e:
        logger.error(f"Error reading document register: {str(e)}")
        raise

def generateForepages(selected_submissions: pd.DataFrame, doc_folder: str):
    # Validate template file
    template_path = 'template.docx'
    if not os.path.isfile(template_path):
        logger.error(f"Template file not found: {template_path}")
        raise FileNotFoundError(f"Template file not found: {template_path}")

    # Validate input DataFrame
    if selected_submissions.empty or not all(col in selected_submissions for col in ['report_title', 'doc_number_pdf', 'submission_date']):
        logger.error("Invalid or empty selected_submissions DataFrame")
        raise ValueError("Selected submissions DataFrame is empty or missing required columns")

    # Validate doc_folder
    if not os.path.isdir(doc_folder) or not os.access(doc_folder, os.W_OK):
        logger.error(f"Invalid or non-writable doc_folder: {doc_folder}")
        raise ValueError(f"Invalid or non-writable folder: {doc_folder}")

    logger.info(f"Using folder for forepages: {doc_folder}")
    
    generated_files = []
    try:
        template = DocxTemplate(template_path)
        
        # Generate Word documents with unique filenames
        timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        for idx, (_, submission) in enumerate(selected_submissions.iterrows()):
            context = {
                'report_title': submission['report_title'],
                'doc_number': submission['doc_number'],
                'submission_date': submission['submission_date'].strftime('%d %b %Y')
            }
            template.render(context=context)
            # Use unique filename to avoid conflicts
            output_filename = f"{submission['doc_number']}_{timestamp}_{idx}.docx"
            output_path = os.path.join(doc_folder, output_filename)
            template.save(output_path)
            generated_files.append(output_path)
            logger.info(f"Generated forepage: {output_path}")
        
        # Create a ZIP file
        zip_filename = f"forepages_{timestamp}.zip"
        zip_path = os.path.join(doc_folder, zip_filename)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                zipf.write(file_path, os.path.basename(file_path))
        logger.info(f"Created ZIP file: {zip_path}")
        
        # Read ZIP file as binary
        with open(zip_path, 'rb') as f:
            zip_data = f.read()
        
        return zip_data, zip_filename
    
    except Exception as e:
        logger.error(f"Error in generateForepages: {str(e)}")
        raise
    
    finally:
        # Clean up generated .docx files (leave .zip and document register)
        for file_path in generated_files:
            try:
                os.remove(file_path)
                logger.info(f"Cleaned up forepage: {file_path}")
            except Exception as e:
                logger.error(f"Error cleaning up forepage {file_path}: {str(e)}")