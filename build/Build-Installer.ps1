<#
.SYNOPSIS
    Builds the add-ins and packages them into the x86 and x64 MSIs.

.DESCRIPTION
    Produces a single Macro_Polo.msi that serves both 32-bit and 64-bit Office.

    Office reads its add-in registration from the registry view matching its own architecture, so
    the package writes both views rather than shipping one MSI per bitness. The payload assemblies
    are AnyCPU and shared; only the registry entries are duplicated. The package itself is x64 and
    so requires 64-bit Windows - which is not the same question as Office's bitness.

.PARAMETER Configuration
    Build configuration. Defaults to Release.

.PARAMETER OutputDirectory
    Where to copy the finished MSIs. Defaults to "artifacts" at the repository root.

.EXAMPLE
    .\build\Build-Installer.ps1
#>
[CmdletBinding()]
param(
    [string] $Configuration = 'Release',
    [string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $repoRoot 'artifacts'
}

function Find-MSBuild {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path $vswhere)) {
        throw 'vswhere.exe not found. MSBuild is required to build the add-in projects.'
    }

    $installation = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
    if (-not $installation) {
        throw 'No Visual Studio installation with MSBuild was found.'
    }

    $msbuild = Join-Path $installation 'MSBuild\Current\Bin\MSBuild.exe'
    if (-not (Test-Path $msbuild)) {
        throw "MSBuild not found at $msbuild."
    }

    return $msbuild
}

$msbuild = Find-MSBuild
Write-Host "Using $msbuild"

# The add-ins are non-SDK-style projects, so they are built with MSBuild rather than `dotnet build`.
foreach ($project in 'Macro_Polo_Word\Macro_Polo_Word.csproj', 'Macro_Polo_Excel\Macro_Polo_Excel.csproj') {
    Write-Host "`nBuilding $project"
    & $msbuild (Join-Path $repoRoot $project) -t:Restore,Build -p:Configuration=$Configuration -v:m -nologo
    if ($LASTEXITCODE -ne 0) { throw "Failed to build $project" }
}

# Gate packaging on the COM surface being intact. Nothing here is visible to the compiler and a
# defect in it takes the whole Office process down, so it is checked before anything ships rather
# than discovered by opening Word.
Write-Host "`nChecking the COM surface"
& (Join-Path $PSScriptRoot 'Test-ComSurface.ps1') -Configuration $Configuration
if ($LASTEXITCODE -ne 0) { throw 'COM surface check failed; refusing to package.' }

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$installerProject = Join-Path $repoRoot 'Macro_Polo_Installer\Macro_Polo_Installer.wixproj'

Write-Host "`nPackaging"
& dotnet build $installerProject -c $Configuration -p:Platform=x64 --nologo -v:m
if ($LASTEXITCODE -ne 0) { throw 'Failed to package the installer' }

$msi = Join-Path $repoRoot "Macro_Polo_Installer\bin\x64\$Configuration\Macro_Polo.msi"
Copy-Item -LiteralPath $msi -Destination $OutputDirectory -Force

Write-Host "`nInstallers written to $OutputDirectory"
Get-ChildItem $OutputDirectory -Filter *.msi | Select-Object Name, Length
