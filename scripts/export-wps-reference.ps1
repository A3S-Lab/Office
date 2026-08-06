param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$WpsPath,

    [ValidateRange(1, 100000)]
    [int]$ExpectedPageCount,

    [ValidateRange(1, 120)]
    [int]$StartupTimeoutSeconds = 15
)

$ErrorActionPreference = 'Stop'
$hasExpectedPageCount = $PSBoundParameters.ContainsKey('ExpectedPageCount')

function Resolve-WpsExecutable {
    param([string]$RequestedPath)

    if ($RequestedPath) {
        $resolved = [IO.Path]::GetFullPath($RequestedPath)
        if (-not [IO.File]::Exists($resolved)) {
            throw "WPS executable does not exist: $resolved"
        }
        return $resolved
    }

    $installationRoot = Join-Path $env:LOCALAPPDATA 'Kingsoft\WPS Office'
    if (-not [IO.Directory]::Exists($installationRoot)) {
        throw "WPS installation root does not exist: $installationRoot"
    }

    $candidates = @(
        Get-ChildItem -LiteralPath $installationRoot -Directory |
            ForEach-Object {
                $candidate = Join-Path $_.FullName 'office6\wps.exe'
                if ([IO.File]::Exists($candidate)) {
                    [pscustomobject]@{
                        Path = $candidate
                        Version = try { [version]$_.Name } catch { [version]'0.0' }
                    }
                }
            } |
            Sort-Object Version -Descending
    )
    if ($candidates.Count -eq 0) {
        throw "No WPS Writer executable was found below $installationRoot"
    }
    return $candidates[0].Path
}

function Get-OwnedWpsProcessIds {
    param(
        [int]$RootProcessId,
        [string]$ExecutablePath
    )

    $processes = @(
        Get-CimInstance -ClassName Win32_Process -Filter "Name = 'wps.exe'" `
            -ErrorAction SilentlyContinue
    )
    $owned = [Collections.Generic.HashSet[int]]::new()
    [void]$owned.Add($RootProcessId)
    do {
        $changed = $false
        foreach ($process in $processes) {
            if (
                $owned.Contains([int]$process.ParentProcessId) -and
                [string]::Equals(
                    [string]$process.ExecutablePath,
                    $ExecutablePath,
                    [StringComparison]::OrdinalIgnoreCase
                ) -and
                $owned.Add([int]$process.ProcessId)
            ) {
                $changed = $true
            }
        }
    } while ($changed)

    return @(
        $processes |
            Where-Object { $owned.Contains([int]$_.ProcessId) } |
            Select-Object -ExpandProperty ProcessId
    )
}

function Invoke-In32BitPowerShell {
    $powershell32 = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if (-not [IO.File]::Exists($powershell32)) {
        throw "32-bit Windows PowerShell does not exist: $powershell32"
    }

    $arguments = @(
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        $PSCommandPath,
        '-InputPath',
        $InputPath,
        '-OutputPath',
        $OutputPath,
        '-StartupTimeoutSeconds',
        [string]$StartupTimeoutSeconds
    )
    if ($WpsPath) {
        $arguments += @('-WpsPath', $WpsPath)
    }
    if ($hasExpectedPageCount) {
        $arguments += @('-ExpectedPageCount', [string]$ExpectedPageCount)
    }

    & $powershell32 @arguments
    exit $LASTEXITCODE
}

if ([IntPtr]::Size -ne 4) {
    Invoke-In32BitPowerShell
}

$resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
$resolvedOutput = [IO.Path]::GetFullPath($OutputPath)
$outputDirectory = [IO.Path]::GetDirectoryName($resolvedOutput)
if (-not $outputDirectory) {
    throw "The output path must include a directory: $resolvedOutput"
}
[IO.Directory]::CreateDirectory($outputDirectory) | Out-Null

$resolvedWps = Resolve-WpsExecutable -RequestedPath $WpsPath
$temporaryOutput = Join-Path $outputDirectory (
    '{0}.{1}.pdf' -f [IO.Path]::GetFileNameWithoutExtension($resolvedOutput), [guid]::NewGuid().ToString('N')
)
$automationClsid = [guid]'{000209FF-0000-4b30-A977-D214852036FF}'
$server = $null
$application = $null
$document = $null
$result = $null

try {
    $server = Start-Process `
        -FilePath $resolvedWps `
        -ArgumentList '/prometheus', '/wps', '/Automation' `
        -WindowStyle Hidden `
        -PassThru

    $deadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
    $lastActivationError = $null
    while ($null -eq $application -and [DateTime]::UtcNow -lt $deadline) {
        try {
            $type = [Type]::GetTypeFromCLSID($automationClsid, $true)
            $application = [Activator]::CreateInstance($type)
        } catch {
            $lastActivationError = $_.Exception.Message
            Start-Sleep -Milliseconds 250
        }
    }
    if ($null -eq $application) {
        throw "WPS automation did not become ready: $lastActivationError"
    }

    $application.Visible = $false
    $application.DisplayAlerts = 0
    $document = $application.Documents.Open($resolvedInput, $false, $true)
    $pageCount = [int]$document.ComputeStatistics(2)
    if (
        $hasExpectedPageCount -and
        $pageCount -ne $ExpectedPageCount
    ) {
        throw "WPS rendered $pageCount pages; expected $ExpectedPageCount."
    }

    $document.ExportAsFixedFormat($temporaryOutput, 17)
    if (-not [IO.File]::Exists($temporaryOutput)) {
        throw "WPS did not create the expected PDF: $temporaryOutput"
    }
    if ((Get-Item -LiteralPath $temporaryOutput).Length -le 0) {
        throw "WPS created an empty PDF: $temporaryOutput"
    }

    Move-Item -LiteralPath $temporaryOutput -Destination $resolvedOutput -Force
    $wpsVersion = (Get-Item -LiteralPath $resolvedWps).VersionInfo.FileVersion
    $result = [pscustomobject]@{
        input = $resolvedInput
        inputSha256 = (Get-FileHash -LiteralPath $resolvedInput -Algorithm SHA256).Hash.ToLowerInvariant()
        output = $resolvedOutput
        outputSha256 = (Get-FileHash -LiteralPath $resolvedOutput -Algorithm SHA256).Hash.ToLowerInvariant()
        pageCount = $pageCount
        wpsExecutable = $resolvedWps
        wpsFileVersion = $wpsVersion
        wpsAutomationVersion = [string]$application.Version
    }
} finally {
    if ($null -ne $document) {
        try {
            $document.Close(0)
        } catch {
            # WPS can close its automation server with the final document.
        }
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        } catch {
            # The RPC object may already be disconnected.
        }
    }
    if ($null -ne $application) {
        try {
            $application.Quit(0)
        } catch {
            # The RPC server may already have exited with the final document.
        }
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
        } catch {
            # The RPC object may already be disconnected.
        }
    }
    if ($null -ne $server) {
        foreach ($processId in Get-OwnedWpsProcessIds -RootProcessId $server.Id -ExecutablePath $resolvedWps) {
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
    }
    if ([IO.File]::Exists($temporaryOutput)) {
        [IO.File]::Delete($temporaryOutput)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$result | ConvertTo-Json
