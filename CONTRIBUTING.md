# Contributing

Thanks for contributing to `izblock`.

## Workflow

1. Create a branch from `main`.
2. Keep changes focused and small.
3. Build locally before opening a PR.
4. Open a PR with a clear summary and test notes.

## Local build checks

Matcher:

```powershell
dotnet build .\apps\iz-matcher\IzMatcher.App\IzMatcher.App.csproj -c Release
```

Archicad placer:

```powershell
dotnet build .\apps\iz-archicad-placer\IzArchicadPlacer.App\IzArchicadPlacer.App.csproj -c Release
```

AutoCAD plugin (requires local AutoCAD API DLLs):

```powershell
dotnet build .\IzAutoCADPlugin\IzAutoCADPlugin.csproj -c Release
```

## Coding notes

1. Keep data schema changes explicit in README/docs.
2. Avoid committing build outputs (`bin`, `obj`, `publish`).
3. Prefer backward-compatible changes in Excel field names when possible.
