$ErrorActionPreference = "Stop"

# Bootstrap + run the AS9102 FAI GUI.
# - Creates/repairs .venv (venvs are machine-specific; do not copy between PCs)
# - Installs app dependencies
# - Launches `python -m as9102_fai`

Set-Location -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) | Out-Null
Set-Location -Path ".." | Out-Null

function Select-Python {
    $candidates = @(
        @{ cmd = "py"; args = @("-3.12") },
        @{ cmd = "py"; args = @("-3.13") },
        @{ cmd = "python"; args = @() }
    )

    foreach ($c in $candidates) {
        if (-not (Get-Command $c.cmd -ErrorAction SilentlyContinue)) {
            continue
        }
        try {
            & $c.cmd @($c.args + @("-c", "import sys; print(sys.executable)")) | Out-Null
            if ($LASTEXITCODE -eq 0) {
                return $c
            }
        } catch {
            continue
        }
    }

    throw "Python was not found. Install Python 3.12+ (or ensure the 'py' launcher is available)."
}

function Test-VenvPython {
    param(
        [Parameter(Mandatory=$true)][string]$VenvPython
    )

    if (-not (Test-Path -Path $VenvPython)) {
        return $false
    }

    try {
        & $VenvPython -c "import sys; print(sys.version)" | Out-Null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

$py = Select-Python
$venvPython = Join-Path -Path (Get-Location) -ChildPath ".venv\Scripts\python.exe"

if (-not (Test-VenvPython -VenvPython $venvPython)) {
    Write-Host "(Re)creating .venv..."
    if (Test-Path -Path ".venv") {
        Remove-Item -Recurse -Force ".venv"
    }

    & $py.cmd @($py.args + @("-m", "venv", ".venv"))

    if (-not (Test-VenvPython -VenvPython $venvPython)) {
        throw "Virtual environment creation failed (expected $venvPython)."
    }

    & $venvPython -m ensurepip --upgrade | Out-Null
    & $venvPython -m pip install --upgrade pip
}

Write-Host "Installing dependencies..."
& $venvPython -m pip install -r "as9102_fai/requirements.txt"

Write-Host "Launching AS9102 FAI..."
& $venvPython -m as9102_fai
