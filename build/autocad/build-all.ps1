$ErrorActionPreference = "Stop"

$singleScript = Join-Path $PSScriptRoot "build-version.ps1"
$years = @("2023", "2024", "2025")

foreach ($year in $years) {
    $autoCADDir = "C:\Program Files\Autodesk\AutoCAD $year"
    if (-not (Test-Path -LiteralPath (Join-Path $autoCADDir "AcMgd.dll"))) {
        Write-Warning "Skipping AutoCAD $year: API DLLs not found in $autoCADDir"
        continue
    }

    & $singleScript -Year $year -AutoCADDir $autoCADDir
}

Write-Host "Done."
