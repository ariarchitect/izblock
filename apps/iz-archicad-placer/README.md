# iz-archicad-placer

Standalone C# app for placing Archicad library elements from Excel mapping data.

## Architecture

- UI: WinForms (`.NET 8`, Windows)
- Source format: Excel (`.xlsx` / `.xls`)
- Integration: Tapir commands against Archicad 26 (JSON-RPC style endpoint)

## Implemented workflow (ported from `temp.py`)

1. Load Excel workbook with multiple sheets.
2. Build table `filename -> floors`.
3. Build table `sheet -> library part name`.
4. Fetch stories from Archicad (`GetStories`).
5. Convert user floors:
   - positive: `1 -> 0`, `2 -> 1`
   - negative: unchanged
6. Place objects via Tapir (`CreateObjects`) using scaled `minx/miny` and story `level`.
7. Write `text1/text2` into properties (`SetPropertyValuesOfElements`) or write `text1` into Element ID.

## Input expectations per sheet

Required columns:

- `filename`
- `minx`
- `miny`
- `text1`

Optional:

- `text2`

## Run

```powershell
dotnet run --project .\IzArchicadPlacer.App\IzArchicadPlacer.App.csproj
```

## Build EXE

```powershell
dotnet publish .\IzArchicadPlacer.App\IzArchicadPlacer.App.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true -o .\publish-win64
```

## Notes

- Default Tapir URL is `http://127.0.0.1:19723`.
- If your Archicad/Tapir endpoint differs, set it in app UI before `Reload Stories`.
