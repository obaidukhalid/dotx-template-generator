<#
    build_windows_exe.ps1
    ---------------------
    Builds dist\TemplateStudio.exe, a single Windows executable that runs
    Template Studio with no Python installed on the target machine.

    Run it from this folder in PowerShell:

        .\build_windows_exe.ps1

    It needs Python on PATH (or the py launcher) on the machine doing the
    build. The machine that *runs* the exe needs nothing at all.

    The build happens in an isolated virtual environment under .build-venv, so
    PyInstaller and its dependencies are never installed into your system
    Python and never touch the .venv used for development.
#>

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$venv   = Join-Path $PSScriptRoot ".build-venv"
$python = Join-Path $venv "Scripts\python.exe"

function Find-Python {
    foreach ($candidate in @("py", "python", "python3")) {
        $found = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($found) { return $found.Source }
    }
    throw "No Python found on PATH. Install it from python.org and tick 'Add Python to PATH'."
}

if (-not (Test-Path $python)) {
    $host_python = Find-Python
    Write-Host "Creating the build environment with $host_python ..." -ForegroundColor Cyan
    & $host_python -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw "Could not create the virtual environment." }
}

Write-Host "Installing the build dependencies ..." -ForegroundColor Cyan
& $python -m pip install --quiet --upgrade pip
& $python -m pip install --quiet "pyinstaller>=6.0" "Flask>=3.0" "python-docx>=1.1"
if ($LASTEXITCODE -ne 0) { throw "Could not install the build dependencies." }

# A stale build/ makes PyInstaller reuse analysis that no longer matches the
# source, which shows up as a missing module at runtime rather than at build
# time. Cheap to avoid.
Write-Host "Clearing previous build output ..." -ForegroundColor Cyan
foreach ($dir in @("build", "dist")) {
    if (Test-Path $dir) { Remove-Item -Recurse -Force $dir }
}

Write-Host "Building TemplateStudio.exe ..." -ForegroundColor Cyan
& $python -m PyInstaller --noconfirm --clean TemplateStudio.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed. The output above says why." }

$exe = Join-Path $PSScriptRoot "dist\TemplateStudio.exe"
if (-not (Test-Path $exe)) { throw "The build reported success but $exe is not there." }

$size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host ""
Write-Host "Done. $exe ($size MB)" -ForegroundColor Green
Write-Host ""
Write-Host "Double click it, or run it from a terminal. It starts the server and"
Write-Host "opens http://127.0.0.1:5000 in your browser. config.json, presets.json"
Write-Host "and the feedback folder are written next to the exe, so put it in a"
Write-Host "folder you can write to rather than straight in Program Files."
