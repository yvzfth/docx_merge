Param()

$ErrorActionPreference = 'Stop'

# Build a Windows .exe using PyInstaller
# Output: dist/SummaryCollector/SummaryCollector.exe

Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location ..

$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) {
  Write-Error 'python not found. Install Python 3.10+ first.'
}

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

Remove-Item -Recurse -Force build, dist, SummaryCollector.spec -ErrorAction SilentlyContinue

python -m PyInstaller `
  --windowed `
  --name "SummaryCollector" `
  gui.py

Write-Host "`nBuilt: dist/SummaryCollector/SummaryCollector.exe"

