# FNC 3 Email Reminder Automation

Generates a personalized reminder for each participant and sends it through an SMTP server. The message lists the participant's outstanding courses and includes the learner handbook, the company signature, and bilingual confidentiality notices.

The repository does not include the company's signature images. The email currently uses the supplied text signature; provide the approved image files to add them inline.

## Input

Use an Excel workbook (`.xlsx`/`.xlsm`) or a UTF-8 CSV with one participant per row and columns for:
Excel input requires `openpyxl` (`pip install openpyxl`); CSV input uses only the Python standard library.

| Field | Recognized headers | Example |
|---|---|---|
| Name | `Nome`, `Name`, `Allievo`, `Nome allievo` | Mario Rossi |
| Email | `Email`, `E-mail`, `Mail`, `Indirizzo email` | mario@example.com |
| Outstanding courses | `Corsi`, `Courses`, `Corsi da finire`, `Corsi da completare`, `Video` | Corso A \| Corso B |

In the courses cell, separate course names with a vertical bar, semicolon, or line break. The script never modifies the source file. If your workbook uses other header names, pass them with `--name-column`, `--email-column` and `--courses-column`.

## Preview and send

Preview is the default; it validates all rows and sends nothing:

```bash
py email_reminder_automation.py "partecipanti.xlsx"
```

Configure SMTP in the environment, then add `--send` to enable delivery. It will display the recipient count and still require typing `INVIA` before connecting and sending:

```powershell
$env:SMTP_HOST = "smtp.office365.com"
$env:SMTP_PORT = "587"
$env:SMTP_USER = "account@azienda.it"
$env:SMTP_PASSWORD = "password-o-app-password"
$env:EMAIL_FROM = "account@azienda.it"
py email_reminder_automation.py "partecipanti.xlsx" --send
```

The script uses STARTTLS and does not save credentials or participant data. Keep the source workbook and credentials on your computer; do not commit either to the repository. Check your organization's SMTP policy and account permissions before sending.

Without valid SMTP configuration, or if input rows contain missing names, invalid or duplicate email addresses, or empty course lists, sending is blocked.
