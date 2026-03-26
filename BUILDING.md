# Multi-Version Build (AutoCAD 2023/2024/2025)

## Build one version

```powershell
.\build\autocad\build-version.ps1 -Year 2024
```

Optional custom AutoCAD path:

```powershell
.\build\autocad\build-version.ps1 -Year 2024 -AutoCADDir "D:\Autodesk\AutoCAD 2024"
```

## Build all installed versions

```powershell
.\build\autocad\build-all.ps1
```

## Package GitHub Release files

Build/publish everything first:

```powershell
.\build\publish-all.ps1
```

Then create GitHub-ready zip archives:

```powershell
.\build\package-github-release.ps1
```

Output is written to:

- `artifacts\github-release\iz-matcher-win64.zip`
- `artifacts\github-release\iz-archicad-placer-win64.zip`
- `artifacts\github-release\IzAutoCADPlugin.bundle.zip`

The script checks:

- `C:\Program Files\Autodesk\AutoCAD 2023`
- `C:\Program Files\Autodesk\AutoCAD 2024`
- `C:\Program Files\Autodesk\AutoCAD 2025`

And skips years that are not installed.

## Bundle layout

Build output is copied into:

- `build\autocad\template\IzAutoCADPlugin.bundle\Contents\2023`
- `build\autocad\template\IzAutoCADPlugin.bundle\Contents\2024`
- `build\autocad\template\IzAutoCADPlugin.bundle\Contents\2025`

`PackageContents.xml` already maps each folder to its AutoCAD series.
