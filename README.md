# Invoice Document Status Automation

A Windows desktop automation that refreshes Excel data, checks transport-document status in a browser portal, applies business rules, and produces a classified spreadsheet.

> Public source edition. Configure credentials locally and use empty or synthetic inputs. Company and institution names identify the original integration context; this repository does not claim affiliation or endorsement.

## Workflow

1. Refresh and read the input workbook with Excel automation.
2. Use Selenium to consult the authorized portal.
3. Apply the status and date rules in Python.
4. Export the resulting classifications to Excel.

## Configuration

Copy `.env.example` to `.env` and fill in the authorized portal URL and account locally before running the script. The committed source no longer contains account credentials.

## Stack and requirements

Python, Pandas, Selenium, OpenPyXL, `pywin32`, Chrome, and `webdriver-manager`. Windows with Microsoft Excel installed is required for the COM refresh steps.

## Run

Install dependencies from `requirements.txt`, prepare the expected workbook and authorized portal configuration, then run `python main.py`. No credentials or customer documents belong in Git.
