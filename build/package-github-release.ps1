$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$repoRoot = Split-Path -Parent $PSScriptRoot
$buildRoot = Join-Path $repoRoot "artifacts\izblock-win64"
$releaseRoot = Join-Path $repoRoot "artifacts\github-release"
$stagingRoot = Join-Path $releaseRoot "_staging"

$packages = @(
    @{
        Type = "exe"
        Source = Join-Path $buildRoot "iz-matcher\IzMatcher.App.exe"
        Archive = Join-Path $releaseRoot "iz-matcher-win64.zip"
    },
    @{
        Type = "exe"
        Source = Join-Path $buildRoot "iz-archicad-placer\IzArchicadPlacer.App.exe"
        Archive = Join-Path $releaseRoot "iz-archicad-placer-win64.zip"
    },
    @{
        Type = "folder"
        Source = Join-Path $buildRoot "IzAutoCADPlugin.bundle"
        Stage = Join-Path $stagingRoot "IzAutoCADPlugin.bundle"
        Archive = Join-Path $releaseRoot "IzAutoCADPlugin.bundle.zip"
    }
)

foreach ($package in $packages) {
    if (-not (Test-Path -LiteralPath $package.Source)) {
        throw "Build output not found: $($package.Source). Run .\build\publish-all.ps1 first."
    }
}

if (Test-Path -LiteralPath $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $releaseRoot | Out-Null
New-Item -ItemType Directory -Force -Path $stagingRoot | Out-Null

foreach ($package in $packages) {
    Write-Host "Preparing $($package.Archive)..." -ForegroundColor Cyan

    if ($package.Type -eq "exe") {
        $archive = [System.IO.Compression.ZipFile]::Open($package.Archive, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive,
                $package.Source,
                [System.IO.Path]::GetFileName($package.Source),
                [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
        finally {
            $archive.Dispose()
        }

        continue
    }

    Copy-Item -LiteralPath $package.Source -Destination $package.Stage -Recurse -Force
    Get-ChildItem -LiteralPath $package.Stage -Recurse -File -Filter "*.pdb" | Remove-Item -Force
    Compress-Archive -Path $package.Stage -DestinationPath $package.Archive -Force
}

Remove-Item -LiteralPath $stagingRoot -Recurse -Force

Write-Host ""
Write-Host "GitHub release packages:" -ForegroundColor Green
Get-ChildItem -LiteralPath $releaseRoot -Filter "*.zip" | ForEach-Object {
    Write-Host " - $($_.FullName)"
}
