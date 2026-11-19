param(
    [string]$PythonExe = "python"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonCommand = Get-Command $PythonExe -ErrorAction SilentlyContinue
if ($null -eq $pythonCommand) {
    throw "Python executable not found: $PythonExe"
}

$pythonPath = $pythonCommand.Source
$buildEnvironment = Join-Path $root "build\venv"
$buildPython = Join-Path $buildEnvironment "Scripts\python.exe"
if (-not (Test-Path $buildPython)) {
    & $pythonPath -m venv $buildEnvironment
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to create the isolated build environment."
    }
}

$requirements = Join-Path $root "requirements.txt"
$buildRequirements = Join-Path $root "requirements-build.txt"
& $buildPython -m pip install --disable-pip-version-check -r $requirements -r $buildRequirements
if ($LASTEXITCODE -ne 0) {
    throw "Failed to install application dependencies in the isolated build environment."
}

$distribution = Join-Path $root "dist"
$work = Join-Path $root "build\pyinstaller"
$specification = Join-Path $root "build\spec"
$entryPoint = Join-Path $root "main.py"
$arguments = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--windowed",
    "--onedir",
    "--name", "ExcelHeatMap",
    "--hidden-import=xlwings._xlwindows",
    "--hidden-import=xlwings.com_server",
    "--hidden-import=xlwings.udfs",
    "--hidden-import=win32timezone",
    "--hidden-import=win32com.client",
    "--hidden-import=win32com.server.dispatcher",
    "--hidden-import=win32com.server.policy",
    "--hidden-import=win32com.server.util",
    "--hidden-import=pythoncom",
    "--hidden-import=pywintypes",
    "--distpath", $distribution,
    "--workpath", $work,
    "--specpath", $specification,
    $entryPoint
)

$pythonBase = & $pythonPath -c "import sys; print(sys.base_prefix)"
if ($LASTEXITCODE -ne 0) {
    throw "Failed to locate the selected Python installation."
}
$runtimeDllDirectory = Join-Path $pythonBase "Library\bin"
if (Test-Path $runtimeDllDirectory) {
    $runtimeDependencyScript = @'
import glob
import os
import pefile
import sys

extension_dir = os.path.join(sys.base_prefix, "DLLs")
runtime_dir = os.path.join(sys.base_prefix, "Library", "bin")
runtime_dlls = set()
for extension in glob.glob(os.path.join(extension_dir, "*.pyd")):
    try:
        image = pefile.PE(extension, fast_load=True)
        image.parse_data_directories(
            directories=[pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_IMPORT"]]
        )
    except (OSError, pefile.PEFormatError):
        continue
    for entry in getattr(image, "DIRECTORY_ENTRY_IMPORT", []):
        name = entry.dll.decode("ascii", errors="replace")
        if name.lower().startswith(("api-ms-win-", "vcruntime")):
            continue
        path = os.path.join(runtime_dir, name)
        if os.path.isfile(path):
            runtime_dlls.add(path)
for path in sorted(runtime_dlls):
    print(path)
'@
    $runtimeDllPaths = $runtimeDependencyScript | & $buildPython -
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to inspect Python extension DLL dependencies."
    }
    foreach ($runtimeDllPath in $runtimeDllPaths) {
        if ($runtimeDllPath) {
            $arguments += "--add-binary=$runtimeDllPath;."
        }
    }
}

Push-Location $root
try {
    & $buildPython @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed with exit code $LASTEXITCODE."
    }

    $applicationFolder = Join-Path $distribution "ExcelHeatMap"
    Copy-Item (Join-Path $root "PORTABLE_README.txt") `
        (Join-Path $applicationFolder "README.txt") -Force
    $archive = Join-Path $distribution "ExcelHeatMap-portable.zip"
    Compress-Archive -Path $applicationFolder -DestinationPath $archive -Force
    Write-Output "Portable application created: $archive"
}
finally {
    Pop-Location
}
