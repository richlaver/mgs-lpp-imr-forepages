# LPP-IMR-FOREPAGES

A Streamlit app to automate the generation of forepages for weekly reports in the Lantau Portfolio Project. The app populates a Word template with data from a Google spreadsheet, streamlining report preparation for the Consultancy Agreement NEX/1110.

## Features

- **Automated Forepage Creation**: Generates Word documents using a template, filling in report title, document number, and submission date.
- **Google Sheets Integration**: Retrieves submission data from a Google spreadsheet document register.
- **Customizable Output**: Allows users to select folders for downloading the document register and saving generated forepages.
- **Date Range Selection**: Filters submissions by a user-defined date range.
- **Intuitive Interface**: Built with Streamlit for a web-based, user-friendly experience.

## Requirements

- Python 3.12 or higher
- Dependencies (see `requirements.txt`):
  - `streamlit==1.38.0`
  - `gdown==5.2.0`
  - `pandas==2.2.2`
  - `docxtpl==0.18.0`
  - `openpyxl==3.1.5`
  - `python-dateutil==2.9.0`

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/<your-username>/LPP-IMR-FOREPAGES.git
   cd LPP-IMR-FOREPAGES
   ```

2. (Optional) Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Start the Streamlit app:
   ```bash
   streamlit run app.py
   ```

2. Open the app in your browser (default: `http://localhost:8501`).

3. Use the interface:
   - **Select Document Register Folder**: Choose where to download the Google spreadsheet.
   - **Scan Document Register**: Downloads and processes the spreadsheet.
   - **Select Submission Period**: Pick a start and end date for submissions.
   - **Select Output Folder**: Specify where to save the generated forepages.
   - **Choose Submissions**: Use checkboxes to select reports for forepage generation.
   - **Generate Forepages**: Creates Word documents in the chosen folder.

## Project Structure

- `app.py`: Main Streamlit application with the user interface.
- `functions.py`: Utility functions for downloading, reading, and generating forepages.
- `template.docx`: Word template for forepage output.
- `requirements.txt`: Python dependencies.
- `.gitignore`: Excludes temporary files (e.g., `.xlsx`, `.docx`, virtual environments).

## Contributing

1. Fork the repository.
2. Create a feature branch:
   ```bash
   git checkout -b feature/your-feature
   ```
3. Commit changes:
   ```bash
   git commit -m "Add your feature"
   ```
4. Push to your fork:
   ```bash
   git push origin feature/your-feature
   ```
5. Open a pull request.

## License

[MIT License](LICENSE) (update with your preferred license).

## Notes

- The Google spreadsheet URL in `functions.py` is publicly accessible. For private spreadsheets, consider using `gspread` with authentication.
- Ensure `template.docx` is in the project root directory.
- For deployment to platforms like Streamlit Cloud, additional configuration may be required (e.g., ensuring `template.docx` is included).

## Contact

For questions or issues, open a GitHub issue or contact the maintainer at [your-email@example.com](mailto:your-email@example.com).