# NominesApp

NominesApp is a Python desktop application for automated payroll distribution.
It takes a multi-page payroll PDF, detects each employee's DNI/NIE using text
extraction or OCR, generates one password-protected PDF per employee, and sends
each payslip by email through SMTP.

## Features

- Split a payroll PDF into one PDF per employee
- Detect DNI/NIE from each page using text extraction or OCR
- Match employees against an Excel file with DNI and email columns
- Protect each generated PDF with the employee's DNI/NIE as password
- Send the generated payslips automatically through SMTP
- Optional debug crop generation for OCR troubleshooting

## Tech Stack

- Python
- PySide6
- PyMuPDF
- pypdf
- pdf2image
- pytesseract
- Poppler
- Tesseract OCR

## Project Structure

- `principal_smtp.py`: main application entry point
- `tools/`: bundled OCR and PDF processing binaries used by the app
- `NominesApp.spec`: PyInstaller build configuration
- `icono.ico`: application icon
- `version.txt`: executable version metadata

## How It Works

1. Select a payroll PDF containing multiple employees' payslips.
2. Select an Excel file with at least two columns: `DNI` and `Email`.
3. Select an output folder.
4. Enter the email subject and body.
5. Choose the month.
6. Generate individual protected PDFs.
7. Send them by SMTP.

Each generated PDF is encrypted using the detected DNI/NIE as its password.

## SMTP Configuration

The application reads SMTP settings from:

`%APPDATA%\NominesApp\smtp_config.json`

If the file does not exist, the app can create a template. A typical config
looks like this:

```json
{
  "smtp_host": "in-v3.mailjet.com",
  "smtp_port": 25,
  "smtp_user": "YOUR_API_KEY",
  "smtp_pass": "YOUR_SECRET_KEY",
  "use_starttls": true,
  "from_email": "payroll@your-domain.com",
  "from_name": "Payroll"
}
```

Do not commit real SMTP credentials to GitHub.

## Running the Project

Run the main script:

```powershell
python principal_smtp.py
```

## Building the Executable

Build with PyInstaller using the provided spec file:

```powershell
pyinstaller NominesApp.spec
```

## Notes

- The project currently depends on the local `tools/` folder for Tesseract and
  Poppler.
- Generated folders such as `build/`, `dist/`, `output/`, and local test data
  should not be committed.
- Payroll PDFs, employee spreadsheets, and SMTP credentials should stay out of
  the repository.

## Repository Recommendation

Recommended files to keep in GitHub:

- `principal_smtp.py`
- `tools/`
- `NominesApp.spec`
- `icono.ico`
- `version.txt`
- `README.md`

## License

No license has been defined yet.
