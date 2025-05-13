import streamlit as st
import gdown
import pandas as pd
import datetime
import os
import tempfile
import zipfile
import shutil
from docxtpl import DocxTemplate
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def downloadDocRegister(output_path: str):
    url = 'https://docs.google.com/spreadsheets/d/1Uvt9HNpA0THgVv42bu8ye-7Zif6VttwyJJ24ywZZk80/export?format=xlsx&id=1Uvt9HNpA0THgVv42bu8ye-7Zif6VttwyJJ24ywZZk80'
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    gdown.download(url, output_path, quiet=False)

def readDocRegister(path: str) -> pd.DataFrame:
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
    return doc_register

def generateForepages(selected_submissions: pd.DataFrame):
    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()
    logger.info(f"Temporary directory created: {temp_dir}")
    
    try:
        template = DocxTemplate('template.docx')
        generated_files = []
        
        # Generate Word documents
        for _, submission in selected_submissions.iterrows():
            context = {
                'report_title': submission['report_title'],
                'doc_number': submission['doc_number'],
                'submission_date': submission['submission_date'].strftime('%d %b %Y')
            }
            template.render(context=context)
            output_path = os.path.join(temp_dir, f"{submission['doc_number']}.docx")
            template.save(output_path)
            generated_files.append(output_path)
            logger.info(f"Generated forepage: {output_path}")
        
        # Create a ZIP file
        zip_filename = f"forepages_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
        zip_path = os.path.join(temp_dir, zip_filename)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                zipf.write(file_path, os.path.basename(file_path))
        logger.info(f"Created ZIP file: {zip_path}")
        
        # Read ZIP file as binary
        with open(zip_path, 'rb') as f:
            zip_data = f.read()
        
        return zip_data, zip_filename
    
    finally:
        # Clean up temporary directory
        try:
            shutil.rmtree(temp_dir)
            logger.info(f"Cleaned up temporary directory: {temp_dir}")
        except Exception as e:
            logger.error(f"Error cleaning up temporary directory: {str(e)}")