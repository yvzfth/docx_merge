# SummaryQuickMerge - Installation Guide

This guide explains how to download and install the **SummaryQuickMerge** application for macOS and Windows.

## What is SummaryQuickMerge?

SummaryQuickMerge is a desktop application that extracts "Summary" sections from multiple Word documents (`.docx` files) and merges them into a single document. It preserves original formatting, fonts, images, and styles.

## System Requirements

- **macOS**: macOS 10.13 (High Sierra) or later
- **Windows**: Windows 10 or later
- **Disk Space**: Approximately 100 MB for the application

## macOS Installation

### Option 1: Pre-built Application (Recommended)

1. **Download** the `SummaryQuickMerge.app` file (or `SummaryQuickMerge.zip` if distributed as a zip archive).

2. **If downloaded as a zip file**, extract it:

   - Double-click the `SummaryQuickMerge.zip` file
   - A `SummaryQuickMerge.app` file will appear in the same folder

3. **First-time launch** (macOS Gatekeeper):

   - macOS may block the app on first launch with a security warning
   - **Control-click** (right-click) on `SummaryQuickMerge.app`
   - Select **"Open"** from the context menu
   - Click **"Open"** in the security dialog
   - The app will be added to your security exceptions and will open normally in the future

4. **Move to Applications** (optional):
   - Drag `SummaryQuickMerge.app` to your `Applications` folder
   - You can now launch it from Launchpad or Spotlight

### Option 2: Build from Source

If you prefer to build the application yourself:

1. Ensure Python 3.10+ and developer tools are installed
2. From the project root:
   ```bash
   chmod +x scripts/build_macos.sh
   scripts/build_macos.sh
   ```
3. The built app will be in `dist/SummaryQuickMerge.app`

### Troubleshooting macOS

- **"App is damaged" error**: This is usually a Gatekeeper issue. Use Control-click → Open as described above.
- **App won't open**: Check that you're running macOS 10.13 or later.
- **Permission denied**: Make sure the app has execute permissions: `chmod +x SummaryQuickMerge.app`

## Windows Installation

### Option 1: Pre-built Application (Recommended)

1. **Download** the `SummaryQuickMerge` folder (or `SummaryQuickMerge.zip` if distributed as a zip archive).

2. **If downloaded as a zip file**, extract it:

   - Right-click `SummaryQuickMerge.zip`
   - Select **"Extract All..."**
   - Choose a destination folder (e.g., `C:\Program Files\` or your desktop)
   - Click **"Extract"**

3. **Navigate** to the extracted `SummaryQuickMerge` folder.

4. **Run the application**:

   - Double-click `SummaryQuickMerge.exe`
   - If Windows SmartScreen shows a warning:
     - Click **"More info"**
     - Click **"Run anyway"**
   - The application window will open

5. **Create a shortcut** (optional):
   - Right-click `SummaryQuickMerge.exe`
   - Select **"Create shortcut"**
   - Drag the shortcut to your desktop or Start menu

### Option 2: Build from Source

If you prefer to build the application yourself:

1. Ensure Python 3.10+ is installed
2. Open PowerShell in the project root and run:
   ```powershell
   PowerShell -ExecutionPolicy Bypass -File scripts\build_windows.ps1
   ```
3. The built executable will be in `dist\SummaryQuickMerge\SummaryQuickMerge.exe`

### Troubleshooting Windows

- **SmartScreen warning**: This is normal for unsigned applications. Click "More info" → "Run anyway" on first launch.
- **Antivirus false positive**: Some antivirus software may flag PyInstaller-built applications. Add an exception if needed.
- **Missing DLL errors**: Ensure you have Windows 10 or later and all Windows updates installed.
- **Permission denied**: Right-click the executable and select "Run as administrator" if needed.

## Using SummaryQuickMerge

1. **Launch** the application (double-click `SummaryQuickMerge.app` on macOS or `SummaryQuickMerge.exe` on Windows).

2. **Select the parent folder** containing your `.docx` files (or subfolders with `.docx` files).

3. **Choose an output file** (defaults to `collected_summaries.docx` in the parent folder).

4. **Optional**: Check "Include .docx files directly in the parent folder" if you want to include files in the root folder, not just subfolders.

5. **Click "Collect Summaries"** and wait for the process to complete.

6. The output file will contain all Summary sections from the scanned documents, with each section labeled with its source filename.

## Uninstallation

### macOS

1. Open Finder
2. Navigate to `Applications` (or wherever you placed the app)
3. Drag `SummaryQuickMerge.app` to the Trash
4. Empty the Trash

### Windows

1. Navigate to the folder where you installed `SummaryQuickMerge`
2. Delete the entire `SummaryQuickMerge` folder
3. Delete any shortcuts you created

## Getting Help

- Check the `README.md` file for detailed usage instructions
- Review the `REPORT.md` file for technical details
- Ensure your `.docx` files contain sections with "Summary" headings

## Security Notes

- SummaryQuickMerge is an open-source application. You can review the source code before building.
- The application only reads `.docx` files you specify and creates a new output file. It does not modify your original files.
- On first launch, macOS and Windows may show security warnings because the application is not code-signed. This is normal for open-source applications.
