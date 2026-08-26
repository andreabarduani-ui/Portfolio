# Certificate Manager

**Phase 3 — Reporting** · Python, pypdf, pandas, xlrd, openpyxl

A four-command pipeline that takes certificates from **bulk print to
company-organized archives**, reading names from the documents themselves.

## Commands

| Command | What it does |
|---|---|
| `dividi` | Splits the bulk-print PDF (front+back pages alternating) into one file per person, reading each name from the certificate text after "Si attesta che" — anything unrecognized still gets saved, nothing is lost |
| `organizza` | Sorts one edition's PDFs into company folders using an Excel directory; name matching is exact first, then fuzzy **within the edition** (avoids homonyms across editions), with a `_DA_VERIFICARE` folder for the leftovers |
| `riorganizza` | Rebuilds multiple organized editions into a company → edition view |
| `tutto` | `dividi` + `organizza` in one command for a single edition |

## Run

```bash
py certificate_manager.py tutto "bulk_print.pdf" "directory.xls" --edizione CORSO10
```

Every argument is optional: with a single candidate file/folder in place, the
tool finds it automatically.

## Value

Hundreds of certificates per edition went from a morning of manual splitting
and filing (with real privacy risk of mis-filed documents) to a single command
with a verifiable audit trail.

> Public showcase copy: company and personal identifiers have been generalized.
