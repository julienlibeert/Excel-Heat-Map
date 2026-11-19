# Rebuild the portable app after changing the code
You’ll need a 64-bit Windows build machine with Python 3.10 or later and internet access to install any missing build dependencies. You don’t need to install them globally—the build script uses its own environment.

Save your code changes and close any running copy of Excel Heat Map.

Open PowerShell and go to the project folder:

``` 
cd "C:\Users\J0305060\OneDrive - TotalEnergies\Documents\Divers\IT\Python\Excel-Heat-Map"
```

Run the build script, specifying the Python executable on your build machine. For example, on the machine used to make the current package:
```
powershell -ExecutionPolicy Bypass -File .\build_portable.ps1 `
  -PythonExe "C:\Users\J0305060\AppData\Local\anaconda3\python.exe"
```

If python is already available in PowerShell, you can instead run:
```
.\build_portable.ps1
```

When the script reports “Portable application created,” find the updated package at:
```
dist\ExcelHeatMap-portable.zip
```

Send or copy that ZIP to the target computer, extract the whole ExcelHeatMap folder, and run ExcelHeatMap.exe inside it. Test the new build before sharing it.

The script installs or updates dependencies in build\venv, rebuilds the executable, and replaces the ZIP. If you change dependencies, update requirements.txt or requirements-build.txt first; the build script installs from those files each time.

The portable app still requires desktop Microsoft Excel on the target computer for workbook access. More details are in PORTABLE_README.txt.