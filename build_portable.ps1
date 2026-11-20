param(
    [string]$PythonExe = "",
    [switch]$RefreshDependencies
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$localBuildRoot = Join-Path $env:LOCALAPPDATA "ExcelHeatMap\build"
New-Item -ItemType Directory -Path $localBuildRoot -Force | Out-Null

function Find-Python {
    if ($PythonExe) {
        $command = Get-Command $PythonExe -ErrorAction Stop
        return $command.Source
    }

    $remembered = Join-Path $localBuildRoot "python-path.txt"
    $candidates = @()
    if (Test-Path $remembered) {
        $candidates += (Get-Content $remembered -Raw).Trim()
    }
    $candidates += Join-Path $root "build\venv\Scripts\python.exe"
    foreach ($name in @("python", "python3")) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($command -and $command.Source -notlike "*\Microsoft\WindowsApps\*") {
            $candidates += $command.Source
        }
    }
    $launcher = Get-Command py -ErrorAction SilentlyContinue
    if ($launcher) {
        $launchedPython = & $launcher.Source -3 -c "import sys; print(sys.executable)"
        if ($LASTEXITCODE -eq 0) {
            $candidates += $launchedPython
        }
    }
    $candidates += @(
        (Join-Path $env:LOCALAPPDATA "anaconda3\python.exe"),
        (Join-Path $env:LOCALAPPDATA "miniconda3\python.exe"),
        (Join-Path $env:USERPROFILE "anaconda3\python.exe"),
        (Join-Path $env:USERPROFILE "miniconda3\python.exe")
    )
    $pythonInstallations = Join-Path $env:LOCALAPPDATA "Programs\Python"
    if (Test-Path $pythonInstallations) {
        $candidates += Get-ChildItem $pythonInstallations -Directory |
            ForEach-Object { Join-Path $_.FullName "python.exe" }
    }
    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if (Test-Path $candidate -PathType Leaf) {
            return $candidate
        }
    }
    throw "No Python installation found. Install 64-bit Python 3.10+ on the build machine, or run with -PythonExe 'C:\path\to\python.exe'. Recipients do not need Python."
}

function Write-BuildLog {
    process {
        $line = "$_"
        Add-Content -LiteralPath $buildLog -Value $line -Encoding UTF8
        Write-Output $line
    }
}

function Invoke-BuildPython {
    param([string[]]$Arguments, [string]$Description)
    Write-Host "`n$Description"
    $previousPreference = $ErrorActionPreference
    try {
        # Windows PowerShell treats native stderr as ErrorRecord objects.
        $ErrorActionPreference = "Continue"
        & $buildPython @Arguments 2>&1 | Write-BuildLog
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousPreference
    }
    if ($exitCode -ne 0) {
        throw "$Description failed (exit $exitCode). See $buildLog"
    }
}

$lock = [System.IO.File]::Open(
    (Join-Path $localBuildRoot "build.lock"),
    [System.IO.FileMode]::OpenOrCreate,
    [System.IO.FileAccess]::ReadWrite,
    [System.IO.FileShare]::None
)
Push-Location $root
try {
    $pythonPath = Find-Python
    $probe = @'
import hashlib, json, platform, struct, sys
if sys.version_info < (3, 10) or struct.calcsize("P") != 8 or sys.platform != "win32":
    raise SystemExit("Building requires 64-bit Windows Python 3.10 or later.")
identity = hashlib.sha256((sys.base_prefix + sys.version).encode()).hexdigest()[:12]
print(json.dumps({
    "base_executable": sys._base_executable,
    "tag": sys.implementation.cache_tag + "-" + platform.machine().lower(),
    "identity": identity
}))
'@
    $probeOutput = $probe | & $pythonPath -
    if ($LASTEXITCODE -ne 0) {
        throw "Cannot use Python at '$pythonPath'. Select a working interpreter with -PythonExe."
    }
    $pythonInfo = $probeOutput | ConvertFrom-Json
    $buildEnvironment = Join-Path $localBuildRoot ("venv-" + $pythonInfo.identity)
    $buildPython = Join-Path $buildEnvironment "Scripts\python.exe"
    $buildLog = Join-Path $localBuildRoot "portable-build.log"
    Set-Content $buildLog "Building ExcelHeatMap with $($pythonInfo.base_executable)" -Encoding UTF8
    if (-not (Test-Path $buildPython)) {
        & $pythonInfo.base_executable -m venv $buildEnvironment
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to create the isolated build environment at $buildEnvironment."
        }
    }
    Set-Content (Join-Path $localBuildRoot "python-path.txt") $pythonInfo.base_executable

    $cache = Join-Path $root ("build\dependencies\" + $pythonInfo.tag)
    $wheels = Join-Path $cache "wheels"
    $runtime = Join-Path $cache ("runtime-" + $pythonInfo.identity)
    New-Item -ItemType Directory -Path $wheels -Force | Out-Null
    $requirements = Join-Path $root "requirements.txt"
    $buildRequirements = Join-Path $root "requirements-build.txt"
    $fingerprint = (Get-FileHash $requirements).Hash + (Get-FileHash $buildRequirements).Hash
    $cacheStamp = Join-Path $cache "requirements.sha256"
    $cachedFingerprint = ""
    if (Test-Path $cacheStamp) {
        $cachedFingerprint = (Get-Content $cacheStamp -Raw).Trim()
    }
    if ($RefreshDependencies -or $cachedFingerprint -ne $fingerprint -or
        -not (Get-ChildItem $wheels -Filter "*.whl" | Select-Object -First 1)) {
        $cacheReady = $false
        if (-not $RefreshDependencies) {
            Write-Host "`nChecking whether the local wheel cache covers the requirements..."
            $previousPreference = $ErrorActionPreference
            try {
                $ErrorActionPreference = "Continue"
                & $buildPython -m pip download --disable-pip-version-check --no-index `
                    --find-links $wheels --dest $wheels -r $requirements -r $buildRequirements 2>&1 |
                    Write-BuildLog
                $cacheReady = $LASTEXITCODE -eq 0
            }
            finally {
                $ErrorActionPreference = $previousPreference
            }
        }
        if (-not $cacheReady) {
            Write-Host "The cache needs additional wheels; contacting the configured package index."
            Invoke-BuildPython -Description "Caching application and build dependency wheels" -Arguments @(
                "-m", "pip", "download", "--disable-pip-version-check", "--only-binary=:all:",
                "--timeout", "15", "--retries", "1",
                "--dest", $wheels, "-r", $requirements, "-r", $buildRequirements
            )
        }
        Set-Content $cacheStamp $fingerprint
    }
    Invoke-BuildPython -Description "Installing dependencies from the reusable local cache" -Arguments @(
        "-m", "pip", "install", "--disable-pip-version-check", "--no-index",
        "--find-links", $wheels, "--upgrade", "-r", $requirements, "-r", $buildRequirements
    )
    Invoke-BuildPython -Description "Checking build dependencies" -Arguments @("-m", "pip", "check")
    Invoke-BuildPython -Description "Collecting Python runtime DLLs, including transitive dependencies" -Arguments @(
        "-m", "excel_heat_map.utils.portable_build", $runtime
    )

    $staging = Join-Path $localBuildRoot "dist"
    $arguments = @(
        "-m", "PyInstaller", "--noconfirm", "--clean", "--windowed", "--onedir",
        "--noupx", "--name", "ExcelHeatMap",
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
        "--distpath", $staging,
        "--workpath", (Join-Path $localBuildRoot "pyinstaller"),
        "--specpath", (Join-Path $localBuildRoot "spec")
    )
    $runtimeManifest = Get-Content (Join-Path $runtime "manifest.json") -Raw | ConvertFrom-Json
    foreach ($name in $runtimeManifest.files) {
        $arguments += "--add-binary=$(Join-Path $runtime $name);."
    }
    $arguments += Join-Path $root "main.py"
    Invoke-BuildPython -Description "Building the portable Windows application" -Arguments $arguments

    $applicationFolder = Join-Path $staging "ExcelHeatMap"
    $executable = Join-Path $applicationFolder "ExcelHeatMap.exe"
    if (-not (Test-Path $executable)) {
        throw "PyInstaller did not create the expected executable: $executable"
    }
    $report = Join-Path $localBuildRoot "self-test.json"
    if (Test-Path $report) { Remove-Item $report -Force }
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $executable
    $startInfo.Arguments = "--self-test `"$report`""
    $startInfo.WorkingDirectory = $applicationFolder
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.EnvironmentVariables["PATH"] = "$env:SystemRoot\System32;$env:SystemRoot"
    foreach ($name in @("PYTHONHOME", "PYTHONPATH", "TCL_LIBRARY", "TK_LIBRARY", "TCLLIBPATH")) {
        $startInfo.EnvironmentVariables.Remove($name)
    }
    $startInfo.EnvironmentVariables["PYTHONNOUSERSITE"] = "1"
    Write-Host "`nTesting the executable without Python or Conda on PATH..."
    $process = [System.Diagnostics.Process]::Start($startInfo)
    try {
        if (-not $process.WaitForExit(90000)) {
            Stop-Process -Id $process.Id -Force
            throw "The executable self-test timed out. See $buildLog and $report"
        }
        if ($process.ExitCode -ne 0 -or -not (Test-Path $report)) {
            throw "The executable self-test failed (exit $($process.ExitCode)). See $report and $buildLog"
        }
    }
    finally {
        $process.Dispose()
    }
    $result = Get-Content $report -Raw | ConvertFrom-Json
    if (-not $result.ok -or -not $result.frozen) {
        throw "The packaged application failed its standalone self-test. See $report"
    }

    $instructions = Join-Path $root "Creating Portable Executable Readme.md"
    Copy-Item $instructions (Join-Path $applicationFolder "README.md") -Force
    $distribution = Join-Path $root "dist"
    New-Item -ItemType Directory -Path $distribution -Force | Out-Null
    $archive = Join-Path $distribution "ExcelHeatMap-portable.zip"
    $pendingArchive = Join-Path $distribution "ExcelHeatMap-portable.building.zip"
    Compress-Archive -Path $applicationFolder -DestinationPath $pendingArchive -Force
    $publishedFolder = Join-Path $distribution "ExcelHeatMap"
    if (Test-Path $publishedFolder) {
        Remove-Item -LiteralPath $publishedFolder -Recurse -Force
    }
    Copy-Item $applicationFolder $distribution -Recurse -Force
    Move-Item $pendingArchive $archive -Force
    Write-Host "`nPortable executable: $(Join-Path $publishedFolder 'ExcelHeatMap.exe')"
    Write-Host "Portable ZIP: $archive"
    Write-Host "Reusable dependency cache: $cache"
    Write-Host "Build log: $buildLog"
    Write-Host "Standalone self-test: $report"
}
finally {
    Pop-Location
    $lock.Dispose()
}
