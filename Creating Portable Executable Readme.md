# Build and use the portable Windows application

## Rebuild after changing Python code

Save your code changes and close any running copy of the portable application.
From the project folder, run:

```powershell
.\build_portable.ps1
```

The build machine needs 64-bit Windows and Python 3.10 or later. The script reuses
a previously selected interpreter, including the existing project build
environment or an Anaconda installation. To select a different interpreter:

```powershell
.\build_portable.ps1 -PythonExe "C:\path\to\python.exe"
```

If PowerShell blocks scripts, use a process-only execution-policy override:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_portable.ps1
```

The first build downloads dependency wheels. Later code-only builds reuse
`build\dependencies\<Python-and-architecture>\wheels` without accessing a package
index. Python/Conda runtime DLLs and their transitive dependencies are automatically
gathered into a neighboring `runtime-<interpreter-id>` folder and included in every
build. Keep `build\dependencies` for future rebuilds; no manual DLL copying is
necessary. Do not copy these cached files into the executable's folder yourself.

The isolated environment, PyInstaller working files, build log, and self-test
report live in `%LOCALAPPDATA%\ExcelHeatMap\build`, outside OneDrive. The script
remembers its Python selection there. Changing either requirements file refreshes
the wheel cache automatically. To deliberately refresh versions within the
requirements' ranges:

```powershell
.\build_portable.ps1 -RefreshDependencies
```

Before publishing, the script runs the frozen executable with Python/Conda
removed from its search path. This hidden self-test checks Windows COM imports,
the Tk interface, 2D/3D scatterplots, trend plotting, and image export without
connecting to Excel or changing any workbooks. A failed self-test stops the build
and writes details to `%LOCALAPPDATA%\ExcelHeatMap\build\self-test.json`.

Successful builds create:

- `dist\ExcelHeatMap\ExcelHeatMap.exe` and its `_internal` runtime folder
- `dist\ExcelHeatMap-portable.zip`, containing the whole portable application

## Run on another computer

Copy the ZIP to a Windows computer, extract the **entire** `ExcelHeatMap` folder,
then run `ExcelHeatMap.exe`. Keep `_internal` alongside the executable.

The recipient does **not** need Python, VS Code, a Python environment, or an
installer. Desktop Microsoft Excel is still required for workbook access; open
the workbook in Excel before selecting it in the application. This is a portable
folder distribution, not a single-file executable.

The self-test checks runtime bundling, not Excel installation, company security
policies, code signing, or antivirus approval. Test the ZIP on a representative
target computer before distributing it broadly.
