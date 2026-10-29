# Chase Statement Converter and Check Payee OCR

A Python command-line prototype for extracting transactions from supported
Chase checking-account statement PDFs and exporting them to CSV.

The tool also attempts to read payee names from embedded check images and
match them to transactions by check number. It assists statement preparation
and manual review; it does not replace accounting judgment.

## Sample Output

The image below shows a spreadsheet preview of the exported transaction data.

![Sample transaction CSV output](Transaction%20CSV%20sample.png)

The preview contains six columns. The current export schema also includes
the reserved `Balance` column described above.

## Scope

This version targets specific Chase checking-statement layouts.

Transaction extraction requires a PDF with a readable text layer.
Check-image OCR does not provide general OCR for fully scanned statements.

Other banks, Chase credit-card statements, and unsupported layouts are
outside the scope of this version.

## Features

- Process PDF files directly inside a specified directory.
- Extract transaction dates, descriptions, check numbers, and amounts.
- Simplify supported transaction descriptions using cleanup rules.
- Suggest bookkeeping accounts using configurable regular-expression rules.
- Attempt to recognize payees in embedded check images using Tesseract OCR.
- Merge and sort transactions by date, with optional date filtering.
- Export incoming and outgoing amounts in separate columns.
- Optionally save per-statement validation results comparing extracted
  counts, amounts, and balances with detected statement summary fields.

## Repository Structure

```text
chase/
  chase.py
.gitignore
LICENSE
README.md
```

Private statements and their derived CSV files and screenshots are not
intended for public distribution.

## Requirements

- Python 3
- Python packages:
  - pdfminer.six
  - pdfplumber
  - pypdfium2
  - Pillow
- Tesseract OCR installed separately, including English language data,
  for check-image recognition.

The OCR-related Python packages are imported when the script starts,
even when an input contains no check images.

Install the Python dependencies:

```bash
python -m pip install pdfminer.six pdfplumber pypdfium2 Pillow
```

This command does not install the Tesseract executable.

The default Windows Tesseract path is:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

Use `--tesseract` to specify a different executable location.

## Usage

Run commands from the repository root.

Create an `input` directory locally and place supported statement PDFs
inside it. Only PDF files directly inside that directory are processed;
subdirectories are not scanned.

```bash
python chase/chase.py --dir "./input" --output "transactions.csv" --validate-log "validation.txt"
```

The CSV and validation report are written to the specified paths.
Without `--output`, the default CSV filename is:

```text
Extracted_Chase_Activity_Stmt_YYYYMMDD.csv
```

To specify another Tesseract installation, replace the executable path
in this Windows example:

```bash
python chase/chase.py --dir "./input" --output "transactions.csv" --tesseract "C:\path\to\tesseract.exe"
```

Display all options:

```bash
python chase/chase.py --help
```

### Command-Line Options

| Option | Purpose |
| --- | --- |
| `--dir` | Directory containing input PDFs. |
| `--output` | Output CSV filename. |
| `--validate-log` | Optional validation-report filename. |
| `--start` | Inclusive start date. |
| `--end` | Exclusive end date. |
| `--date-format` | Date format for date arguments and CSV output. Default: `%m/%d/%Y`. |
| `--account-rules` | Optional CSV containing `Pattern,Account` rules. |
| `--tesseract` | Path to the Tesseract executable. |

Run without date filters when comparing extracted totals against a
complete statement's summary. Filtered transactions may not match
full-statement totals.

## Account Categorization

The tool includes built-in rules for suggesting bookkeeping accounts.

A custom rule file can use the following structure:

```csv
Pattern,Account
DEMO OFFICE STORE,Office Expense
DEMO TRANSFER,Transfer
```

These are illustrative rules, not extracted financial records or tax advice.

Supply your rule file using `--account-rules`. Rules are checked in order,
and the first matching rule provides the account label.

Unmatched transactions and check transactions remain unclassified.
All account suggestions require human review.

## Output

The current CSV schema is:

```text
Trans Date,Check No,Description,Account,Amount,Out Amount,Balance
```

- `Trans Date`: Parsed transaction date.
- `Check No`: Check number where identified.
- `Description`: Cleaned description, potentially replaced with an
  OCR-extracted check payee.
- `Account`: Suggested bookkeeping account, or blank if unclassified.
- `Amount`: Incoming amount.
- `Out Amount`: Outgoing amount represented as a positive value.
- `Balance`: Reserved column, currently left blank. It is not a calculated
  running balance.

When `--validate-log` is provided, a separate validation report is saved.
Checks depend on successfully detecting the relevant statement summary
fields.

A generated CSV or successful process exit does not establish that every
transaction was extracted correctly. Review the output and validation
results before use.

## Known Limitations

- Parsing is layout-dependent; not every Chase statement is supported.
- Fully scanned statements are not supported by the transaction parser.
- Check OCR relies on detectable image-section markers and check-number
  labels in the PDF.
- Handwriting and low-quality images can produce incorrect payee names.
- OCR confidence and low-confidence warnings appear in the console,
  not in the CSV or validation report.
- Low-confidence OCR results may still be written to the description field.
- OCR failures or a missing Tesseract executable can stop processing.
  There is no skip-OCR option or guaranteed fallback.
- Validation mismatches do not prevent CSV export.
- Overlapping statements are not automatically deduplicated.
- Cross-statement continuity checking is not implemented in this version.
- No automated test suite or reproducible sample-PDF test set is included
  in this version.

## Privacy and Demonstration Data

Bank statements and their derived CSV files and screenshots may contain
sensitive information, including account references, payees, dates,
and transaction amounts.

Do not publish real customer statements or transaction outputs.
Use clearly labeled synthetic data for public demonstrations, or complete
a privacy review before sharing redacted material.

Review screenshots separately from their source files. Do not include
credentials, account details, or private generated outputs in the repository.

## My Contributions

Building on the existing MIT-licensed project, I adapted and extended
the workflow in four areas:

1. **Rule-based account categorization**
   Added rules that map supported non-check transaction descriptions
   to bookkeeping account labels in the `Account` column.
   Unmatched transactions remain blank for manual review.

2. **Transaction description normalization**
   Added cleanup rules to simplify descriptions and remove selected
   banking metadata, making exported records easier to review.

3. **Separate incoming and outgoing amount columns**
   Changed the export to place incoming amounts in `Amount` and outgoing
   amounts, represented as positive values, in `Out Amount`.

4. **Check payee recognition**
   Added Tesseract-based OCR to extract payee information from embedded
   check images and match it to transactions using check numbers.

These adaptations were developed with technical guidance and feedback.

For the case-study submission, prior work, case-specific changes,
AI assistance, verification performed, and actual time spent are disclosed
separately.

## Attribution and License

This project adapts existing MIT-licensed work.

The original copyright notice, including
Copyright (c) 2017 siddweiker, is retained in LICENSE.

See [LICENSE](LICENSE) for the license text.

This project is not affiliated with or endorsed by Chase.
