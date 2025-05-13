import streamlit as st
import gdown
import pandas as pd
import datetime
import os
from docxtpl import DocxTemplate

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

def generateForepages(selected_submissions: pd.DataFrame, output_folder: str):
    os.makedirs(output_folder, exist_ok=True)
    template = DocxTemplate('template.docx')
    for _, submission in selected_submissions.iterrows():
        context = {
            'report_title': submission['report_title'],
            'doc_number': submission['doc_number'],
            'submission_date': submission['submission_date'].strftime('%d %b %Y')
        }
        template.render(context=context)
        output_path = os.path.join(output_folder, f"{submission['doc_number']}.docx")
        template.save(output_path)