$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
if (-not (Test-Path ".venv")) {
  py -m venv .venv
}
& .\.venv\Scripts\Activate.ps1
python -m pip install --quiet -r requirements.txt
python run.py
