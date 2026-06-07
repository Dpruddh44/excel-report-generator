# Automation Tool

This tool provides a streamlined web dashboard to automate the generation of operational reports by combining, filtering, and standardizing data across three specific Excel datasets.

## Features
- **Simple Dashboard Interface:** Built with Streamlit for a seamless file-upload experience.
- **Automated Data Processing:** Handles extraction, L4 filtering, deduplication, and ID standardization automatically without manual intervention.
- **Dynamic Template Matching:** Automatically identifies and carries over dynamic columns from your operational templates.
- **One-Click Generation:** Upload your three datasets, click "Generate Report", and instantly download the compiled final report.

## Requirements
Ensure you have the required dependencies installed:
```bash
pip install pandas openpyxl streamlit
```

## How to Run
1. Navigate to the folder containing the project files.
2. Start the Streamlit server:
   ```bash
   streamlit run app.py
   ```
3. The dashboard will automatically open in your web browser.
4. Upload the three required Excel files into their respective upload fields.
5. Click **Generate Report**.
6. Once the processing is complete, a download button will appear allowing you to save the generated report.
