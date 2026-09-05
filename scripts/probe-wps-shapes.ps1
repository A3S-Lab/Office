param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$WpsPath,

    [switch]$IncludeConnector
)

$ErrorActionPreference = 'Stop'

function Resolve-WpsExecutable {
    param([string]$RequestedPath)

    if ($RequestedPath) {
        $resolved = [IO.Path]::GetFullPath($RequestedPath)
        if (-not [IO.File]::Exists($resolved)) {
            throw "WPS executable does not exist: $resolved"
        }
        return $resolved
    }

    $root = Join-Path $env:LOCALAPPDATA 'Kingsoft\WPS Office'
    $candidate = Get-ChildItem -LiteralPath $root -Directory -ErrorAction Stop |
        ForEach-Object { Join-Path $_.FullName 'office6\wps.exe' } |
        Where-Object { [IO.File]::Exists($_) } |
        Select-Object -First 1
    if (-not $candidate) {
        throw "No WPS Writer executable was found below $root"
    }
    return $candidate
}

function Invoke-In32BitPowerShell {
    $powershell32 = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if ([IntPtr]::Size -ne 4) {
        if (-not [IO.File]::Exists($powershell32)) {
            throw "32-bit Windows PowerShell does not exist: $powershell32"
        }
        $arguments = @(
            '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', $PSCommandPath, '-OutputPath', $OutputPath
        )
        if ($WpsPath) { $arguments += @('-WpsPath', $WpsPath) }
        if ($IncludeConnector) { $arguments += '-IncludeConnector' }
        & $powershell32 @arguments
        exit $LASTEXITCODE
    }
}

Invoke-In32BitPowerShell

$resolvedOutput = [IO.Path]::GetFullPath($OutputPath)
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($resolvedOutput)) | Out-Null
$resolvedWps = Resolve-WpsExecutable $WpsPath
$automationClsid = [guid]'{000209FF-0000-4b30-A977-D214852036FF}'
$server = $null
$application = $null
$document = $null

try {
    $server = Start-Process -FilePath $resolvedWps -ArgumentList '/prometheus', '/wps', '/Automation' -WindowStyle Hidden -PassThru
    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    while ($null -eq $application -and [DateTime]::UtcNow -lt $deadline) {
        try {
            $type = [Type]::GetTypeFromCLSID($automationClsid, $true)
            $application = [Activator]::CreateInstance($type)
        } catch {
            Start-Sleep -Milliseconds 250
        }
    }
    if ($null -eq $application) { throw 'WPS automation did not become ready.' }

    $application.Visible = $false
    $application.DisplayAlerts = 0
    $document = $application.Documents.Add()
    # Keep the native shape records uncluttered so the probe can be used as a
    # deterministic OOXML reference for the browser importer.
    $document.TrackRevisions = $false

    # WPS uses the Word Shapes collection for both AutoShapes and connectors.
    $rectangle = $document.Shapes.AddShape(1, 72, 72, 144, 72)
    $rectangle.Name = 'A3S Rectangle'
    $rectangle.Fill.Visible = -1
    $rectangle.Fill.ForeColor.RGB = [int]0xD9EAD3
    $rectangle.Line.Visible = -1
    $rectangle.Line.ForeColor.RGB = [int]0x4472C4
    $rectangle.TextFrame.TextRange.Text = 'WPS shape'

    if ($IncludeConnector) {
        $connector = $document.Shapes.AddConnector(1, 72, 200, 216, 200)
        $connector.Name = 'A3S Connector'
        $connector.Line.ForeColor.RGB = [int]0xC00000
    }

    $document.SaveAs2($resolvedOutput, 16)
    [pscustomobject]@{
        output = $resolvedOutput
        wps = $resolvedWps
        version = [string]$application.Version
        shapeCount = [int]$document.Shapes.Count
    } | ConvertTo-Json -Compress
} finally {
    if ($null -ne $document) { try { $document.Close(0) } catch {} }
    if ($null -ne $application) { try { $application.Quit(0) } catch {} }
    if ($null -ne $server) {
        Get-Process -Id $server.Id -ErrorAction SilentlyContinue |
            Stop-Process -Force -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
