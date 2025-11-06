## Summary Collector (DOCX)

Collect only the "Summary" section from multiple `.docx` files located in subfolders of a parent folder, and merge them into a single `.docx`. The script preserves original fonts, styles, and images (including image sizes) by trimming source documents in place and merging at the document level.

### Features

- Scans all subfolders under a parent directory for `.docx` files
- Detects the "Summary" section near the beginning or end of each document
- Keeps images and formatting intact
- Labels each Summary heading with the source filename, e.g., `Summary (document1)`
- Saves the merged result inside the parent folder by default

### Requirements

- Python 3.10+
- Libraries: `python-docx`, `docxcompose`, `lxml`, `Pillow`
  - Already available in the included virtual environment at `venv/`

### Install Python 3.10+ (Windows & macOS)

- Windows:

  - Download the latest Python 3.10+ installer from `https://www.python.org/downloads/windows/`.
  - During installation, check "Add Python to PATH".
  - Verify: `python --version` or `py -V`.

- macOS:
  - Option 1 (official): download from `https://www.python.org/downloads/macos/`.
  - Option 2 (Homebrew): `brew install python@3.11` (or newer).
  - Verify: `python3 --version`.

### Create and activate a virtual environment, then install packages

- Windows (PowerShell or CMD):

  ```bash
  python -m venv venv
  venv\\Scripts\\activate
  python -m pip install --upgrade pip
  pip install python-docx docxcompose lxml Pillow
  ```

- macOS (Terminal):
  ```bash
  python3 -m venv venv
  source venv/bin/activate
  python -m pip install --upgrade pip
  pip install python-docx docxcompose lxml Pillow
  ```

### Activate the virtual environment (optional but recommended)

```bash
source ~/Desktop/docx/venv/bin/activate
```

### Usage

Pass the parent folder as a positional argument. By default, the output is written to `<parent>/collected_summaries.docx`.

```bash
~/Desktop/docx/venv/bin/python ~/Desktop/docx/main.py \
  "~/Desktop/docx"
```

Include `.docx` files directly under the parent folder as well (not just subfolders):

```bash
~/Desktop/docx/venv/bin/python ~/Desktop/docx/main.py \
  "~/Desktop/docx" \
  --include-root
```

Override the output path if needed:

```bash
~/Desktop/docx/venv/bin/python ~/Desktop/docx/main.py \
  "~/Desktop/docx" \
  --output "~/Desktop/docx/custom_output.docx"
```

### Graphical User Interface (GUI)

If you prefer a windowed app, launch the GUI:

```bash
~/Desktop/docx/venv/bin/python ~/Desktop/docx/gui.py
```

- Select the parent folder containing subfolders with `.docx` files.
- Optionally check "Include .docx files directly in the parent folder".
- Choose an output file (defaults to `collected_summaries.docx` under the parent).
- Click "Collect Summaries". The app runs in the background and shows status.

### Build Executables (macOS & Windows)

You can package the GUI into a standalone app using PyInstaller. Build on each OS to create native binaries for that OS.

#### macOS

1. Ensure Python 3.10+ and developer tools are installed.
2. From the project root:

```bash
chmod +x scripts/build_macos.sh
scripts/build_macos.sh
```

Output: `dist/SummaryQuickMerge.app`

Notes:

- First run may be blocked by Gatekeeper. Control-click the app → Open.
- To share, zip the `.app` or run a DMG packager if desired.
- The build script prefers `SummaryQuickMerge.spec`, which bundles required data for `python-docx`/`docxcompose` and includes `lxml` hidden imports. If you customize packaging, ensure these are included or the app may fail to merge.

Packaging details:

- `python-docx` ships template/data files that must be present at runtime; we collect them via PyInstaller’s `collect_data_files`.
- `docxcompose` may also require data files; those are likewise collected.
- `lxml` uses compiled C extensions; hidden imports `lxml.etree` and `lxml._elementpath` are explicitly included to avoid runtime import errors in the bundled app.

#### Windows

Open PowerShell in the project root and run:

```powershell
PowerShell -ExecutionPolicy Bypass -File scripts\build_windows.ps1
```

Output: `dist\SummaryQuickMerge\SummaryQuickMerge.exe`

Notes:

- If SmartScreen warns, choose “More info” → “Run anyway”.
- Ship the entire folder under `dist\SummaryQuickMerge` or re-run with `--onefile` (advanced).
- The build script prefers `SummaryQuickMerge.spec`, which bundles required data for `python-docx`/`docxcompose` and includes `lxml` hidden imports.

Packaging details:

- Data files from `python-docx` and `docxcompose` are collected so templates/styles are available in the bundled app.
- Hidden imports for `lxml` (`lxml.etree`, `lxml._elementpath`) are added so XML parsing works in the packaged binary.

### Troubleshooting (packaged app)

- If the app fails during merging, a popup will show the full error and traceback.
- A background log is also written to: `~/Desktop/SummaryQuickMergeLogs/app.log`.
- If you rebuilt the app without the spec file, rebuild using the provided scripts so template data is bundled correctly.

### How it works (high level)

1. Each source document is parsed and its block-level elements (paragraphs, tables) are enumerated.
2. The script identifies the paragraph most likely to be the Summary heading (prefers heading styles and positions near the top/bottom).
3. Content outside the Summary range is removed in-place, which preserves images and relationships.
4. The Summary heading is appended with the source filename (without .docx extension), e.g., ` (document3)`.
5. The section includes all content paragraphs below the Summary heading until the next section header (detected by style, heading level, or font size matching the Summary heading).
6. All trimmed documents are merged using `docxcompose.Composer`, preserving page content, styles, and media.

### Notes

- Temporary Word files (like `~$something.docx`) are ignored.
- If no Summary sections are found, an output file will be created with a short notice.
- The Summary detection is heuristic-based. It prefers heading-styled paragraphs containing "Summary" and includes all content paragraphs below until the next section header (same style, same/higher heading level, or same/larger font size). The next section header is excluded from the Summary section.
