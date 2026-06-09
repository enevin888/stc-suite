# SmartTech SOO Revision Review

Streamlit QA/QC tool for reviewing changes between an original Sequence of Operations (SOO) document and an updated SOO document.

This app is currently focused on SOO revision review only. Points-list review, wiring-diagram review, and full-package review are not part of the active workflow.

## What It Does

- Accepts an original SOO and an updated SOO.
- Extracts readable text from Markdown, text, Word, or PDF files.
- Compares the two documents line by line.
- Displays added, removed, and changed-line counts.
- Generates engineering review comments for common SOO concerns.
- Shows a side-by-side diff with an option to focus only on changed areas.

## Review Checks

The app flags items such as:

- Added alarm references that may need trigger, delay, annunciation, and reset review.
- Removed alarm references that may need confirmation.
- Setpoint, threshold, or adjustable-value wording changes.
- Numbering or formatting changes that may affect sequence clarity.
- Possible contradictions between modulation and fixed-position language.
- Shared output logic that may need multi-unit behavior review.
- Fan, damper, and pressure-control coordination language.
- Freeze or low-temperature protection language.

These checks are rule-based and intended as a first-pass engineering review aid. They do not replace a final engineer review.

## Supported SOO Files

- `.md`
- `.txt`
- `.docx`
- `.pdf`

PDF and Word support depend on the parser packages listed in `requirements.txt`.

## Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
