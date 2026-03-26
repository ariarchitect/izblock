$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$artifactsRoot = Join-Path $repoRoot "artifacts\izblock-win64"

if (Test-Path -LiteralPath $artifactsRoot) {
    Remove-Item -LiteralPath $artifactsRoot -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $artifactsRoot | Out-Null

$appTargets = @(
    @{
        Name = "iz-matcher"
        Project = Join-Path $repoRoot "iz-matcher\IzMatcher.App\IzMatcher.App.csproj"
        Output = Join-Path $artifactsRoot "iz-matcher"
    },
    @{
        Name = "iz-archicad-placer"
        Project = Join-Path $repoRoot "iz-archicad-placer\IzArchicadPlacer.App\IzArchicadPlacer.App.csproj"
        Output = Join-Path $artifactsRoot "iz-archicad-placer"
    }
)

foreach ($t in $appTargets) {
    Write-Host "Publishing $($t.Name)..." -ForegroundColor Cyan
    dotnet publish $t.Project `
        -c Release `
        -r win-x64 `
        --self-contained true `
        /p:PublishSingleFile=true `
        /p:IncludeNativeLibrariesForSelfExtract=true `
        -o $t.Output
}

$bundleSource = Join-Path $repoRoot "build\autocad\template\IzAutoCADPlugin.bundle"
$bundleOutput = Join-Path $artifactsRoot "IzAutoCADPlugin.bundle"
Write-Host "Copying AutoCAD bundle..." -ForegroundColor Cyan
Copy-Item -Recurse -Force $bundleSource $bundleOutput

Write-Host ""
Write-Host "Done. Output folders:" -ForegroundColor Green
foreach ($t in $appTargets) {
    Write-Host " - $($t.Output)"
}
Write-Host " - $bundleOutput"
Write-Host ""
Write-Host "Artifacts root: $artifactsRoot"
