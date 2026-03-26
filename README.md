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

- `iz-matcher/` - desktop app for matching blocks and texts.
- `iz-archicad-placer/` - desktop app for Archicad placement.
- `build/` - build and publish scripts.
- `IzAutoCADPlugin/` - AutoCAD plugin source code.
- `build/autocad/template/IzAutoCADPlugin.bundle/` - AutoCAD bundle template used for staging and release packaging.
- `artifacts/` - local generated release outputs.
- `docs/`, `tests/`, `samples/` - placeholders for specs, validation, and example data.

Planned:

- `iz-matcher/` - app that reads exported Excel and matches blocks + texts.
- `iz-archicad-placer/` - app that places Archicad library elements and writes properties.
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
.\build\autocad\build-version.ps1 -Year 2024
```

Build all installed versions:

```powershell
.\build\autocad\build-all.ps1
```

Detailed notes: [BUILDING.md](BUILDING.md)

Publish desktop apps (`iz-matcher` + `iz-archicad-placer`) as single-file EXE:

```powershell
.\build\publish-all.ps1
```

Package GitHub Release archives:

```powershell
.\build\package-github-release.ps1
```

Build output is collected into:

- `artifacts\izblock-win64\iz-matcher\`
- `artifacts\izblock-win64\iz-archicad-placer\`
- `artifacts\izblock-win64\IzAutoCADPlugin.bundle\`

GitHub-ready `.zip` files are collected into:

- `artifacts\github-release\iz-matcher-win64.zip`
- `artifacts\github-release\iz-archicad-placer-win64.zip`
- `artifacts\github-release\IzAutoCADPlugin.bundle.zip`

The app zip files currently contain only the published `.exe` file for each desktop app.

## Repository tooling

- Solution: `izblock.sln`
- Contribution guide: [CONTRIBUTING.md](CONTRIBUTING.md)
- License: [LICENSE](LICENSE)

## Roadmap

1. Stabilize Excel export schema (versioned format).
2. Implement deterministic block-text matching rules.
3. Build Archicad placement engine with property mapping profiles.
4. Add test datasets and end-to-end validation.

## Notes

- Autodesk API assemblies are not stored in the repository.
- AutoCAD bundle template lives in `build\autocad\template\IzAutoCADPlugin.bundle\`.
- Build outputs and generated bundle binaries are ignored by `.gitignore`.
