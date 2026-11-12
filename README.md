# Word to PDF Sanitizer

A Windows desktop application that exports Word documents to PDF, removes selected
metadata, verifies text and rendered appearance, and archives unchanged originals.

## Requirements

- Windows desktop with Microsoft Word installed and activated.
- Windows PowerShell with WPF support.
- The Codex bundled Python runtime at
  `%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`.
- Python packages: pypdf, pypdfium2, Pillow and lxml. NumPy accelerates bitmap
  comparison when available; Pillow provides the fallback. Test fixtures also use
  python-docx, reportlab and openpyxl.

This release's launchers currently expect that bundled runtime. A standalone Python
installer or executable is not included.

## Run

Extract the release ZIP and double-click `Open Word to PDF.cmd`. Select Starting
and Originals directories, preview the file list, then convert. Read
`Instructions.txt` before processing documents. Keep Word closed during a queue.

## Features

- Recursive .doc/.docx/.docm conversion with persistent Word direct PDF export.
- Configurable comment, tracked-change, hidden-text, property, link and tag policies.
- PDF sanitation, structural checks, all-page text and pixel comparisons.
- Checksum-verified originals, restoration records and protected failure handling.
- Parallel PDF verification and background source staging.
- Excel queue reports, diagnostics and detailed Accept/Revert/Undo review.

Verification does not guarantee the absence of every possible form of sensitive
information. Visible document contents remain the user's responsibility; image
payload metadata is not inspected. Accessibility-tag removal reduces screen-reader
support. See the complete instructions for limitations and option tradeoffs.

## Validation

Version 1.8.0 passed 73 automated tests and interface checks. Matched full-queue
tests converted 16 documents with equivalent text/pixels, archive checksums and
review outcomes. Performance varies by document, machine and storage.

Run `python -m unittest discover -s Tests -p "test_*.py"` with the required test
dependencies. Word integration tests require Word and an interactive Windows desktop.

## Public package contents

Source, launchers, instructions and tests only. User preferences, logs, caches,
conversion reports, restoration records and client documents are excluded.
