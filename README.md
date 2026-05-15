# Resume Updated Weekly Automation

Automated resume synchronization system for OPMS onboarding workers.

This project automatically:

- Retrieves onboarding workers from OPMS
- Downloads latest resume documents
- Uploads resumes to Azure Blob Storage
- Extracts resume text from PDF/DOCX files
- Updates resume cache weekly
- Runs automatically using Azure Functions

---

# Features

## Weekly Automated Resume Sync

Every Monday 2:00 AM Perth Time:

1. Load onboarding workers
2. Search OPMS training records
3. Download latest resume files
4. Upload files to Azure Blob Storage
5. Extract text content
6. Update cache automatically

---

# Technologies

- Python 3.11
- Azure Functions
- Azure Blob Storage
- OPMS API
- Pandas
- PyMuPDF
- Flask
- GitHub

---

# Project Structure

```text
Resume_updated_weekly/
│
├── function_app.py
├── GetResumeUpdated.py
├── Getonbroadingpeople.py
├── Getcompetency.py
├── requirements.txt
├── .env
│
├── resume_downloads/
├── Archived/
└── templates/
