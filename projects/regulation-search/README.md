# Regulation Search

**Cross-cutting utility** · Python, pdfplumber (pypdf fallback)

An interactive search engine **inside funding-body manuals, calls and
guidelines**: finds what a regulation says about a specific topic, with page
number and context.

## What it does

- Indexes any regulation PDF page by page
- Offers **predefined research categories** with tuned keyword sets:
  plan variations, reporting deadlines and documents, eligible costs,
  training modes, state-aid regimes and more
- Supports **free-text search** on top of the categories
- Every hit returns **page number + surrounding context**, so the answer can
  be verified on the source immediately

## Run

```bash
py regulation_search.py path/to/manual.pdf
# or without arguments: the script asks for the path interactively
```

## Value

"What does this fund's manual allow on cost variations?" used to mean
re-reading 100-page PDFs. Now it's a menu-driven lookup that always cites
the page.

> Public showcase copy: company and personal identifiers have been generalized.

## Sample output

An illustrative output example (fictional data) is available in [sample-output.md](sample-output.md).
