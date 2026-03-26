# iz-matcher

Second stage of `izblock`: matching blocks with nearby texts.

## What is implemented

- Windows desktop app on `.NET 8` (WinForms).
- Reads source `result.xlsx` with sheets:
- Reads source Excel with sheets:
  - `blocks`
  - `texts`
  - supported formats: `.xlsx`, `.xls`
- Filter fields:
  - blocks: `block_name`, `layer` (+ regex mode)
  - texts: `text_content`, `layer`, `color`, `height min/max` (+ regex mode)
- Two coordinate windows (`Text1`, `Text2`) relative to each block insertion point.
- Match logic equivalent to notebook:
  - select texts in each window for the same `file_name`
  - sort by `insert_x`
  - concatenate `text_plain`
- Preview table and save to output `.xlsx` sheet with columns:
  - `filename`, `minx`, `miny`, `text1`, `text2`
- App state is stored in `%APPDATA%\IZBLOCK\izmatcher.settings.json`.

## Run from source

```powershell
dotnet run --project .\IzMatcher.App\IzMatcher.App.csproj
```

## Build

Framework-dependent build:

```powershell
dotnet publish .\IzMatcher.App\IzMatcher.App.csproj -c Release -o .\publish
```

Single-file self-contained EXE (recommended for distribution):

```powershell
dotnet publish .\IzMatcher.App\IzMatcher.App.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true -o .\publish-win64
```
