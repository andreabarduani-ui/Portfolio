# Adhesion Letter Compiler

**Phase 1 — Program Launch** · Python, python-docx, openpyxl

Fills the official Word adhesion-letter template from an Excel data sheet,
producing one personalized letter for every trainee–company pair.

## What it does

- Reads the Word template and the Excel sheet (people block + companies block)
- Pairs them positionally ("first with first, second with second")
- Writes each field **centered inside its box**, using tab-stop positions
  computed from the geometry of the original PDF layout
- Saves one letter per company; the template is **never modified**

## Run

```bash
py adhesion_letter_compiler.py
```

Files (template, data, output prefix, paragraph indices) are configurable in
the constants at the top of the script.

## Value

What used to be days of retyping letters for each course edition is now a
two-second batch job with pixel-accurate alignment.

> Public showcase copy: company and personal identifiers have been generalized.
