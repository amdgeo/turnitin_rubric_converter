# Turnitin Rubric Converter

Turnitin rubric files use the `.rbc` extension and contain JSON data that is not
convenient to edit or reuse. This project converts those files to a readable
CSV matrix or a structured Excel workbook.

## Features

- CSV export with UTF-8 encoding and spreadsheet formula-injection protection.
- Excel export with `Criteria Matrix`, `Criteria Details`, `Scale Values`, and
  optional `Metadata` worksheets.
- Validation of required sections, IDs, and criterion/scale references.
- Preview and overwrite confirmation in the desktop application.
- Command-line conversion for scripts and automation.
- Cross-platform GitHub Actions tests and tagged-release packaging.

## Desktop application

Install Python 3.9 or later, then run:

```bash
python -m pip install -r requirements.txt
python rubric_converter_gui.py
```

Select an `.rbc` file, choose CSV or Excel, select an output path, and click
**Preview** or **Convert**. The executable releases are available on the
[Releases page](https://github.com/amdgeo/turnitin_rubric_converter/releases).

### Windows installation

For the easiest setup on Windows, download
`Turnitin-Rubric-Converter-Setup.exe` from the latest release and run it. The
installer adds the app to the Start menu, can optionally add a desktop
shortcut, and includes an uninstaller. The portable `rubric-converter.exe`
download remains available for users who prefer not to install the app.

## Command line

```bash
python rubric_converter_cli.py input.rbc output.csv
python rubric_converter_cli.py input.rbc output.xlsx --format excel
python rubric_converter_cli.py input.rbc output.csv --use-name-and-value
python rubric_converter_cli.py ./rubrics ./converted --batch --format excel
```

The output format is inferred from `.csv` or `.xlsx` when `--format` is not
provided. Batch mode converts every `.rbc` file in a directory. Invalid files
produce a clear error and a non-zero exit status.

## Development

Run the test suite with:

```bash
python -m pip install -r requirements.txt
python -m pytest
```

To create a standalone desktop executable locally:

```bash
python -m pip install pyinstaller
pyinstaller --onefile --name rubric-converter rubric_converter_gui.py
```

Tagged pushes such as `v2026.1` run the cross-platform build and attach the
three executables and the Windows installer to a GitHub release.
