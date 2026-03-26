# izblock

Repository: `https://github.com/ariarchitect/izblock`

`izblock` is a workflow for transferring CAD content from AutoCAD into Archicad through a structured intermediate model.

## Project goal

Create a reliable pipeline:

1. Export block and text data from AutoCAD drawings into Excel.
2. Match blocks and texts into semantic pairs/groups (restore their relationship).
3. Place Archicad library elements by block coordinates and write mapped text values into element properties.

## Repository structure (target)

Current:

- `IzAutoCADPlugin/` - AutoCAD plugin source code.
- `IzAutoCADPlugin.bundle/` - AutoCAD bundle layout for deployment.
- `scripts/` - build scripts for versioned AutoCAD builds.

Planned:

- `apps/iz-matcher/` - app that reads exported Excel and matches blocks + texts.
- `apps/iz-archicad-placer/` - app that places Archicad library elements and writes properties.
- `docs/` - format specs, mapping rules, and integration notes.

## AutoCAD plugin

The plugin exports data from `ModelSpace` into `result.xlsx` with two sheets:

1. `blocks`
2. `texts`

Current supported AutoCAD targets:

1. AutoCAD 2023 (`R24.2`)
2. AutoCAD 2024 (`R24.3`)
3. AutoCAD 2025 (`R25.0`)

## Build

Build one version:

```powershell
.\scripts\build-version.ps1 -Year 2024
```

Build all installed versions:

```powershell
.\scripts\build-all.ps1
```

Detailed notes: [BUILDING.md](BUILDING.md)

## Roadmap

1. Stabilize Excel export schema (versioned format).
2. Implement deterministic block-text matching rules.
3. Build Archicad placement engine with property mapping profiles.
4. Add test datasets and end-to-end validation.

## Notes

- Autodesk API assemblies are not stored in the repository.
- Build outputs and generated bundle binaries are ignored by `.gitignore`.
