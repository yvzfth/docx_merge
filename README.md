## Summary Collector (DOCX)

Collect only the "Summary" section from multiple `.docx` files located in subfolders of a parent folder, and merge them into a single `.docx`. The script preserves original fonts, styles, and images (including image sizes) by trimming source documents in place and merging at the document level.

### Features

- Scans all subfolders under a parent directory for `.docx` files
- Detects the "Summary" section near the beginning or end of each document
- Keeps images and formatting intact
- Labels each Summary heading with the source filename, e.g., `Summary (document1.docx)`
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
source /Users/fatih/Desktop/docx/venv/bin/activate
```

### Usage

Pass the parent folder as a positional argument. By default, the output is written to `<parent>/collected_summaries.docx`.

```bash
/Users/fatih/Desktop/docx/venv/bin/python /Users/fatih/Desktop/docx/main.py \
  "/Users/fatih/Desktop/docx"
```

Include `.docx` files directly under the parent folder as well (not just subfolders):

```bash
/Users/fatih/Desktop/docx/venv/bin/python /Users/fatih/Desktop/docx/main.py \
  "/Users/fatih/Desktop/docx" \
  --include-root
```

Override the output path if needed:

```bash
/Users/fatih/Desktop/docx/venv/bin/python /Users/fatih/Desktop/docx/main.py \
  "/Users/fatih/Desktop/docx" \
  --output "/Users/fatih/Desktop/docx/custom_output.docx"
```

### How it works (high level)

1. Each source document is parsed and its block-level elements (paragraphs, tables) are enumerated.
2. The script identifies the paragraph most likely to be the Summary heading (prefers heading styles and positions near the top/bottom).
3. Content outside the Summary range is removed in-place, which preserves images and relationships.
4. The Summary heading is appended with the source filename, e.g., ` (document3.docx)`.
5. All trimmed documents are merged using `docxcompose.Composer`, preserving page content, styles, and media.

### Notes

- Temporary Word files (like `~$something.docx`) are ignored.
- If no Summary sections are found, an output file will be created with a short notice.
- The Summary detection is heuristic-based. It prefers heading-styled paragraphs containing "Summary" and uses the next heading (or document end) to bound the section.
