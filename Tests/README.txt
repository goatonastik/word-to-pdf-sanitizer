WORD TO PDF SANITIZER - AUTOMATED TESTS

These files are development and release checks. The application does not load them
during normal conversion, so keeping them here adds no conversion overhead.

From the Word to PDF program folder, run the main automated suite with:

  python -m unittest discover -s Tests -p "test_*.py"

Run the interface checks with:

  powershell -File Tests\Test-InterfaceLayout.ps1
  powershell -File Tests\Test-ProcessLauncher.ps1
  powershell -File Tests\Test-LogTail.ps1

The real-Word integration scripts create temporary sample documents. They require
Microsoft Word to be closed before they begin:

  python Tests\test_review_workflow.py
  python Tests\test_word_review_cases.py
  powershell -File Tests\Test-SavedViews.ps1 -Folder <generated-fixture-folder>
