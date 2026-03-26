# Multi-Version Build (AutoCAD 2023/2024/2025)

## Build one version

```powershell
.\scripts\build-version.ps1 -Year 2024
```

Optional custom AutoCAD path:

```powershell
.\scripts\build-version.ps1 -Year 2024 -AutoCADDir "D:\Autodesk\AutoCAD 2024"
```

## Build all installed versions

```powershell
.\scripts\build-all.ps1
```

The script checks:

- `C:\Program Files\Autodesk\AutoCAD 2023`
- `C:\Program Files\Autodesk\AutoCAD 2024`
- `C:\Program Files\Autodesk\AutoCAD 2025`

And skips years that are not installed.

## Bundle layout

Build output is copied into:

- `IzAutoCADPlugin.bundle\Contents\2023`
- `IzAutoCADPlugin.bundle\Contents\2024`
- `IzAutoCADPlugin.bundle\Contents\2025`

`PackageContents.xml` already maps each folder to its AutoCAD series.
