Excel Heat Map and Trend Navigator - Portable Windows App

Run ExcelHeatMap.exe from the extracted ExcelHeatMap folder. The folder is
portable: it does not require Python, VS Code, an installer, or administrator
rights on the target computer. Keep the complete folder together when moving
the application.

This application connects to open workbooks through Microsoft Excel. Desktop
Excel must be installed on the target computer, and the workbook must be open
in Excel. Workbook changes, including saved graph settings, remain in memory
until you save the workbook in Excel.

Build the portable package on 64-bit Windows with Python 3.10 or later:

  powershell -ExecutionPolicy Bypass -File .\build_portable.ps1 -PythonExe "C:\path\to\python.exe"

The build script creates an isolated environment under build\venv and installs
the application and packaging dependencies there; it does not modify the
selected Python installation.

The resulting dist\ExcelHeatMap-portable.zip can be copied to another
compatible Windows computer and extracted to any writable folder.
