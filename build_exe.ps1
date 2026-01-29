$ErrorActionPreference = "Stop"
$python = "python"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptRoot

$entry = @{ Script = "run_menu.py"; Name = "timetable_menu" }

& $python -m PyInstaller --onefile --noconfirm --clean --name $entry.Name --distpath dist --workpath build $entry.Script

Write-Host "Build complete. Menu EXE generated in dist folder."
