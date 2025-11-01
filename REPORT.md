## Project Report: DOCX Summary Collector

This report explains, in plain language, what the application does, how it was designed to preserve images and formatting, and the improvements made along the way. It’s written so that a non-technical reader can understand the value and be confident using the tool.

### What this app does

- **Goal**: Go through a parent folder’s subfolders, find Word `.docx` files, extract only the **Summary** section from each, and merge them into a single Word document.
- **Keeps formatting**: Fonts, text styles, and images (including image sizes) are preserved.
- **Labels each section**: The Summary heading is annotated with its source file name, like `Summary (document9.docx)`.
- **Order you expect**: Files are ordered using a human-friendly (natural) sort, so `document10.docx` correctly comes after `document9.docx`.
- **Simple to run**: You pass the parent folder to scan, and the merged file is saved inside that folder by default.

### Why preserving formatting and images is hard (and how we solved it)

- A naïve approach would read text and rebuild paragraphs/tables in a new document. That often **breaks styles and loses images**.
- Instead, we **trim each source document in place** down to just its Summary section. This keeps the underlying document relationships intact (which is how images, styles, and other formatting are actually connected to content).
- Finally, we **merge entire documents** using a document-level composer, which is purpose-built to keep styles, images, and layout.

### How we detect the “Summary” section

- The app finds the paragraph most likely to be the Summary heading using a small heuristic:
  - **Prefers headings** (e.g., normal Word “Heading” styles).
  - **Checks for the word ‘summary’** (case-insensitive).
  - **Prefers headings near the beginning or end** of the document (common for summaries).
- Once the Summary heading is found, the app includes content **from that heading up to the next heading** (or to the end of the document if there isn’t another heading).

### How merging works without breaking anything

1. Collect `.docx` files from subfolders of the parent folder (optionally including files in the parent folder itself).
2. For each file, **trim the document in place** to only the Summary range.
3. Append the **source filename** to the Summary heading, e.g., `Summary (document3.docx)`.
4. Merge the trimmed documents into a single output using a **document-level composer**. This preserves styles, images, and sizes.

### Getting the order right (natural/human sorting)

- Standard string sorting would place `document10.docx` before `document2.docx` (because it compares character-by-character).
- We added a **natural sort**: numeric parts are compared as numbers, so `document9.docx` comes before `document10.docx`.

### Avoiding unwanted files

- Ignores Word’s temporary files that start with `~$`.
- Skips the output file itself so we don’t accidentally re-ingest it.

### Command-line usage (simple)

- Activate the included virtual environment (optional but recommended):

```bash
source ~/Desktop/docx/venv/bin/activate
```

- Run the tool with the parent folder as a positional argument:

```bash
~/Desktop/docx/venv/bin/python ~/Desktop/docx/main.py \
  "~/Desktop/docx"
```

- The output is saved to: `~/Desktop/docx/collected_summaries.docx`

- Include `.docx` files located directly in the parent folder as well:

```bash
~/Desktop/docx/venv/bin/python ~/Desktop/docx/main.py \
  "~/Desktop/docx" \
  --include-root
```

- Choose a custom output filename (optional):

```bash
~/Desktop/docx/venv/bin/python ~/Desktop/docx/main.py \
  "~/Desktop/docx" \
  --output "~/Desktop/docx/custom_output.docx"
```

### Installation guide (Windows & macOS)

This tool requires Python 3.10 or newer and a few Python packages.

- Install Python 3.10+

  - Windows:
    - Download from `https://www.python.org/downloads/windows/`.
    - During install, check "Add Python to PATH".
    - Verify: `python --version` or `py -V`.
  - macOS:
    - Official: download from `https://www.python.org/downloads/macos/`.
    - Or with Homebrew: `brew install python@3.11` (or newer).
    - Verify: `python3 --version`.

- Create a virtual environment and install dependencies
  - Windows (PowerShell/CMD):
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

### What we built and improved (chronological highlights)

- **Initial collector**: A working scan of subfolders that extracted Summary text.
- **Preservation upgrade**: Switched from copying content programmatically (which can lose styles/images) to trimming documents in place and performing **document-level merging** to **fully preserve** formatting and images.
- **Section labeling**: Appends the source filename to each Summary heading, e.g., `Summary (document1.docx)`.
- **Simpler CLI**: Now takes the **parent folder** as a positional argument and saves the output inside that folder by default.
- **User docs**: Added in-file comments at **critical points** and created a **README** with requirements and usage.
- **Correct ordering**: Implemented **natural sorting** of files so numeric parts sort like humans expect (9 before 10).

### What to expect from the output

- One combined Word document with each Summary section clearly labeled.
- Original look and feel of content is maintained (fonts, styles, images, image sizes).
- If no Summary is detected anywhere, the output document will note that instead of failing.

### Limitations and tips

- The Summary detection uses heuristics; it works best when the Summary is in a paragraph that contains the word “Summary” and ideally uses a heading style.
- If an individual document has no recognizable Summary heading, that file is simply skipped.

### Files of interest

- `main.py`: The main script (scanning, trimming, merging, sorting, labeling).
- `README.md`: Quick start guide and usage instructions.
- `collected_summaries.docx`: The produced combined document (saved inside the parent folder by default).

### Final thoughts

This tool gives you a clean, single Word document of all Summary sections across many files—**without** losing images or formatting. It is robust for everyday use and simple to run, while handling tricky details (like images and sort order) so you don’t have to.
