## Add metadata to AutoCAD Plant 3D elements

This add-in applies JSON-defined metadata to Plant 3D elements using a Plant project `Project.xml`.

## Build the add-in (dotnet)

Prereqs:
- AutoCAD Plant 3D 2026 installed.
- AutoCAD Plant 3D 2026 SDK installed (for `PnP*.dll` references).
- .NET SDK 8.x installed.

The project hard-codes SDK and install paths in `addin/MetadataApplier.csproj`:
- `PlantSdkDir=C:\Users\aleja\Documents\AutoCAD Plant 3D 2026 SDK\inc-x64\`
- `AcadDir=C:\Program Files\Autodesk\AutoCAD 2026\`
- `PlantDir=C:\Program Files\Autodesk\AutoCAD 2026\PLNT3D\`

Build (Release):
```powershell
dotnet build .\addin\MetadataApplier.csproj -c Release
```

Output DLL:
`addin\bin\Release\net8.0-windows\MetadataApplier.dll`

## Run the add-in via the Python helper

The helper runs AutoCAD Plant 3D, loads the DLL, and executes the command
`P3D_APPLY_JSON_METADATA_XML`. It passes inputs via environment variables.

Default usage (uses the defaults hard-coded in `run_addin.py`):
```powershell
python .\run_addin.py
```

Note: if you rely on defaults, replace the `PLANT_PROJECT_XML` default path in
`run_addin.py` with your own Project.xml location.

Override paths with environment variables:
```powershell
$env:PLANT_PROJECT_XML = "C:\path\to\Project.xml"
$env:PLANT_JSON_IN = "C:\path\to\metadata.json"
$env:PLANT_PLUGIN_DLL = "C:\path\to\MetadataApplier.dll"
python .\run_addin.py
```

Optional env vars used by the helper:
- `PLANT_SAVE_CHANGES` set to `0` or `1` (default is `1`).
- `PLANT_LOG_PATH` to control the JSONL log file location.

Other defaults inside `run_addin.py`:
- `acad.exe`: `C:\Program Files\Autodesk\AutoCAD 2026\acad.exe`
- Working dir: `C:\PlantAutomationRun` (creates `run_rev2.scr` + log files)
