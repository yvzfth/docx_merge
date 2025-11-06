Param()

$ErrorActionPreference = 'Stop'

# Build a Windows .exe using PyInstaller
# Output: dist/SummaryQuickMerge/SummaryQuickMerge.exe

Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location ..

$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) {
  Write-Error 'python not found. Install Python 3.10+ first.'
}

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

Remove-Item -Recurse -Force build, dist, SummaryQuickMerge.spec -ErrorAction SilentlyContinue

# Build (prefer spec if present for correct data collection)
if (Test-Path ".\SummaryQuickMerge.spec") {
  python -m PyInstaller SummaryQuickMerge.spec
}
else {
  python -m PyInstaller `
    --windowed `
    --name "SummaryQuickMerge" `
    --collect-data docx `
    --collect-data docxcompose `
    --hidden-import lxml.etree `
    --hidden-import lxml._elementpath `
    gui.py
}

Write-Host "`nBuilt: dist/SummaryQuickMerge/SummaryQuickMerge.exe"

