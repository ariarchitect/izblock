param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("2023", "2024", "2025")]
    [string]$Year,

    [string]$AutoCADDir
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$bundleRoot = Join-Path $repoRoot "build\autocad\template\IzAutoCADPlugin.bundle"
$projectPath = Join-Path $repoRoot "IzAutoCADPlugin\IzAutoCADPlugin.csproj"
$bundleTarget = Join-Path $bundleRoot "Contents\$Year"

if ([string]::IsNullOrWhiteSpace($AutoCADDir)) {
    $AutoCADDir = "C:\Program Files\Autodesk\AutoCAD $Year"
}

$requiredDlls = @("AcMgd.dll", "AcDbMgd.dll", "AcCoreMgd.dll", "AdWindows.dll")
foreach ($dll in $requiredDlls) {
    $path = Join-Path $AutoCADDir $dll
    if (-not (Test-Path -LiteralPath $path)) {
        throw "AutoCAD API DLL not found: $path"
    }
}

$outputPath = Join-Path $repoRoot "IzAutoCADPlugin\bin\Release\$Year\net48\"

dotnet build $projectPath -c Release `
    -p:AutoCADDir="$AutoCADDir" `
    -p:OutputPath="$outputPath"

New-Item -ItemType Directory -Force -Path $bundleTarget | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $bundleTarget "Resources") | Out-Null

Copy-Item -Force (Join-Path $outputPath "*.dll") $bundleTarget
Copy-Item -Force (Join-Path $outputPath "Resources\*.png") (Join-Path $bundleTarget "Resources")

Write-Host "Built AutoCAD $Year and copied artifacts to $bundleTarget"
