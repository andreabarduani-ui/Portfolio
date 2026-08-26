# Student Import Generator

**Phase 1 — Program Launch** · Python, pandas, openpyxl

Converts the student data-collection spreadsheet into the platform-ready
import file, automatically enriching each record by parsing the Italian tax
ID (codice fiscale).

## What it does

- Reads the raw "student data request" Excel (tax ID, first name, last name…)
- Parses the tax ID to extract: **gender**, **birth town** and **birth province**
- Resolves town codes through an **embedded national municipality database** (Belfiore)
- Writes the import spreadsheet in the exact format the course platform expects
- CLI with `argparse`: course ID, company VAT number, site, intermediary, ATECO code

## Run

```bash
py student_import_generator.py students.xlsx --id-corso 12345 --piva-azienda 01234567890
```

## Value

Manual enrollment prep required retyping every field and looking up birth
places by hand. The tool cuts it to a single command and eliminates typos on
hundreds of rows per course edition.

> Public showcase copy: company and personal identifiers have been generalized.
