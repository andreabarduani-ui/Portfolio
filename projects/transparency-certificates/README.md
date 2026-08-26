# Transparency Certificates

**Phase 3 — Reporting** · Python, python-docx, openpyxl

Generates the end-of-course **"acquired skills" certificates** — the official
transparency document certifying each participant's learning outcomes.

## What it does

- Reads the Word template (student data, training experience, competencies,
  assessment tables) and the master Excel with one row per participant
- Fills **only the empty value cells** of the template tables:
  institution and responsible-person data are never touched
- Saves one certificate per participant, named by surname, into
  `Attestati_Generati/`
- The template itself is **never modified** — every certificate is a copy

## Run

```bash
py transparency_certificates.py
```

## Value

Dozens of certificates per course used to be filled one by one from a roster.
Now: one command, zero transcription errors, consistent formatting.

> Public showcase copy: company and personal identifiers have been generalized.

## Sample output

An illustrative output example (fictional data) is available in [sample-output.md](sample-output.md).
