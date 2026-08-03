<#
.SYNOPSIS
    Renders the status banner to PNG files so its layout can be inspected without opening Office.

.DESCRIPTION
    The banner is hosted by Office as an ActiveX control in a task pane, which makes visual
    problems awkward to iterate on: every change otherwise means rebuild, reinstall, restart Word.
    This draws the real control offscreen at a range of widths and states and writes PNGs.

    Widths matter because the text rewraps, which changes the height the pane needs. Each state
    matters because the wording differs a lot in length between them.

.PARAMETER Configuration
    Which build output to render. Defaults to Release.

.PARAMETER OutputDirectory
    Where to write the PNGs. Defaults to "artifacts\banner".

.EXAMPLE
    powershell -File build\Render-Banner.ps1
#>
[CmdletBinding()]
param(
    [string] $Configuration = 'Release',
    [string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repoRoot 'artifacts\banner' }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$core = Join-Path $repoRoot "Macro_Polo.Core\bin\$Configuration\net472\Macro_Polo.Core.dll"
if (-not (Test-Path $core)) { throw "Not built: $core" }
$assembly = [System.Reflection.Assembly]::LoadFrom($core)

$evaluator = $assembly.GetType('Macro_Polo.Core.MacroStatusEvaluator', $true)
$presenter = $assembly.GetType('Macro_Polo.Core.MacroStatusPresenter', $true)
$docType   = $assembly.GetType('Macro_Polo.Core.DocumentMacroInfo', $true)
$setType   = $assembly.GetType('Macro_Polo.Core.MacroSecuritySettings', $true)
$levelType = $assembly.GetType('Macro_Polo.Core.VbaWarningLevel', $true)
$bannerType= $assembly.GetType('Macro_Polo.Core.MacroStatusBanner', $true)

# One case per visual state, since the wording length differs sharply between them.
$cases = @(
    @{ Name = 'no-macros';        HasVba = $false; Signed = $false; Level = 'DisableWithNotification'; Trust = 'Unknown' }
    @{ Name = 'runs-unsigned';    HasVba = $true;  Signed = $false; Level = 'EnableAll';               Trust = 'Unknown' }
    @{ Name = 'runs-signed';      HasVba = $true;  Signed = $true;  Level = 'EnableAll';               Trust = 'Unknown' }
    @{ Name = 'needs-consent';    HasVba = $true;  Signed = $false; Level = 'DisableWithNotification'; Trust = 'Unknown' }
    @{ Name = 'needs-publisher';  HasVba = $true;  Signed = $true;  Level = 'DisableExceptSigned';     Trust = 'NotTrusted' }
    @{ Name = 'blocked-unsigned'; HasVba = $true;  Signed = $false; Level = 'DisableExceptSigned';     Trust = 'Unknown' }
    # The state that used to be reported as "blocked until you allow them": already run, because
    # the signer is a publisher this machine trusts.
    @{ Name = 'runs-trusted';     HasVba = $true;  Signed = $true;  Level = 'DisableWithNotification'; Trust = 'Trusted' }
)

# Narrow is a split-screen Word window; wide is a maximised one on a large display.
$widths = 420, 760, 1400

# The readme shows one width, chosen to match a typical document window.
$readmeWidth = 760
$readmeDirectory = Join-Path $repoRoot 'images'

foreach ($case in $cases) {
    $doc = [Activator]::CreateInstance($docType)
    $doc.HasVbaProject = $case.HasVba
    $doc.IsVbaSigned = $case.Signed
    $doc.FullPath = 'C:\Users\someone\Documents\quarterly report.docm'

    $trustType = $assembly.GetType('Macro_Polo.Core.PublisherTrust', $true)
    $sigType = $assembly.GetType('Macro_Polo.Core.VbaSignature', $true)
    $reason = if ($case.Trust -eq 'NotTrusted') { 'the certificate is not in Trusted Publishers' } else { $null }
    $doc.Signature = [Activator]::CreateInstance($sigType, @(
        [Enum]::Parse($trustType, $case.Trust), 'Contoso Ltd', 'DC45AA80E3F3709B0B1C54AC0DF3D7E2403C0348', $reason))

    $settings = [Activator]::CreateInstance($setType)
    $settings.WarningLevel = [Enum]::Parse($levelType, $case.Level)

    $status = $evaluator.GetMethod('Evaluate').Invoke($null, @($doc, $settings))
    $view = $presenter.GetMethod('Describe').Invoke($null, @($status))

    foreach ($width in $widths) {
        $banner = [Activator]::CreateInstance($bannerType)
        try {
            $banner.Update($view)

            # Match what the host does: measure at the pane width, then apply that height.
            $height = $banner.GetPreferredPaneHeight($width)
            $banner.Size = New-Object System.Drawing.Size($width, $height)
            $banner.PerformLayout()

            # Re-measure now the control actually has its size, as the resize handler would.
            $height = $banner.GetPreferredPaneHeight($width)
            $banner.Size = New-Object System.Drawing.Size($width, $height)
            $banner.PerformLayout()

            $bitmap = New-Object System.Drawing.Bitmap($width, $height)
            $banner.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $width, $height)))

            $file = Join-Path $OutputDirectory ("{0}-{1}.png" -f $case.Name, $width)
            $bitmap.Save($file, [System.Drawing.Imaging.ImageFormat]::Png)

            # The readme's screenshots are generated from the real control, so they cannot drift
            # out of date the way hand-captured ones did.
            if ($width -eq $readmeWidth) {
                $bitmap.Save((Join-Path $readmeDirectory ("banner-{0}.png" -f $case.Name)), [System.Drawing.Imaging.ImageFormat]::Png)
            }

            $bitmap.Dispose()

            "{0,-18} width={1,5}  height={2,4}" -f $case.Name, $width, $height
        }
        finally {
            $banner.Dispose()
        }
    }
}

"`nPNGs written to $OutputDirectory"
