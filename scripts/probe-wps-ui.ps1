param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [ValidateSet('shell', 'fields', 'all')]
    [string]$Profile = 'shell',

    [string]$WpsPath,

    [switch]$Json
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
    param([string]$ScriptPath)

    $powershell32 = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if ([IntPtr]::Size -ne 4) {
        if (-not [IO.File]::Exists($powershell32)) {
            throw "32-bit Windows PowerShell does not exist: $powershell32"
        }
        $arguments = @(
            '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', $ScriptPath, '-OutputPath', $OutputPath, '-Profile', $Profile
        )
        if ($WpsPath) { $arguments += @('-WpsPath', $WpsPath) }
        if ($Json) { $arguments += '-Json' }
        & $powershell32 @arguments
        exit $LASTEXITCODE
    }
}

function Get-ControlRecord {
    param($Control)

    try {
        return [pscustomobject]@{
            caption = [string]$Control.Caption
            id = [int]$Control.Id
            type = [int]$Control.Type
            visible = [bool]$Control.Visible
            enabled = [bool]$Control.Enabled
        }
    } catch {
        return [pscustomobject]@{ error = $_.Exception.Message }
    }
}

function Get-BarRecord {
    param(
        $Bar,
        [bool]$IncludeControls
    )

    $controls = @()
    $controlError = $null
    if ($IncludeControls) {
        try {
            $controls = @($Bar.Controls | ForEach-Object { Get-ControlRecord $_ })
        } catch {
            $controlError = $_.Exception.Message
        }
    }

    $record = [pscustomobject]@{
        name = [string]$Bar.Name
        index = [int]$Bar.Index
        visible = [bool]$Bar.Visible
        position = [int]$Bar.Position
        rowIndex = [int]$Bar.RowIndex
        controlCount = @($controls).Count
        controls = $controls
    }
    if ($controlError) { $record | Add-Member -NotePropertyName controlError -NotePropertyValue $controlError }
    return $record
}

$fieldCommandIds = @(
    125, 126, 154, 155, 156, 157, 159, 160, 163, 215, 2384, 3214,
    5731, 5908, 5909, 6069, 6345, 6346, 6347, 6348, 6693, 7116,
    743, 1554, 13464, 3000046, 30077
)

function Test-FieldCommand {
    param($Control)

    return $fieldCommandIds -contains [int]$Control.Id
}

Invoke-In32BitPowerShell $PSCommandPath

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
    $window = $application.ActiveWindow
    $selection = $application.Selection

    $allBars = @($application.CommandBars)
    $fieldBars = @('Header and Footer', 'Mail Merge', 'Fields', 'Form Fields', 'Contents Popup Menu', 'NavigationPane Document Field Popup', 'Ribbon', 'Status Bar')
    $bars = switch ($Profile) {
        'all' { @($allBars | ForEach-Object { Get-BarRecord $_ $true }); break }
        'fields' {
            @($allBars | Where-Object { $fieldBars -contains [string]$_.Name } | ForEach-Object { Get-BarRecord $_ $true })
            break
        }
        default {
            @($allBars | Where-Object { @('Ribbon', 'Status Bar') -contains [string]$_.Name } | ForEach-Object { Get-BarRecord $_ $true })
            break
        }
    }

    $fieldCommands = @(
        foreach ($bar in @($bars)) {
            foreach ($control in @($bar.controls)) {
                $caption = [string]$control.caption
                if (-not (Test-FieldCommand $control)) { continue }
                [pscustomobject]@{
                    bar = [string]$bar.name
                    caption = $caption
                    id = [int]$control.id
                    type = [int]$control.type
                    visible = [bool]$control.visible
                    enabled = [bool]$control.enabled
                }
            }
        }
    )

    $payload = [pscustomobject]@{
        schemaVersion = 1
        probe = 'wps-writer-ui'
        profile = $Profile
        wps = [pscustomobject]@{
            executable = $resolvedWps
            version = [string]$application.Version
            build = [string]$application.Build
            caption = [string]$application.Caption
            automationClsid = $automationClsid.ToString()
        }
        window = [pscustomobject]@{
            caption = [string]$window.Caption
            viewType = [int]$window.View.Type
            windowState = [int]$window.WindowState
            bounds = [pscustomobject]@{
                top = [int]$window.Top
                left = [int]$window.Left
                width = [int]$window.Width
                height = [int]$window.Height
            }
        }
        selection = [pscustomobject]@{
            type = [int]$selection.Type
            start = [int]$selection.Start
            end = [int]$selection.End
            textLength = ([string]$selection.Text).Length
        }
        commandBars = [pscustomobject]@{
            total = @($allBars).Count
            selected = @($bars)
        }
        fieldCommands = $fieldCommands
    }
    $jsonText = $payload | ConvertTo-Json -Depth 8 -Compress
    Set-Content -LiteralPath $resolvedOutput -Value $jsonText -Encoding UTF8
    if ($Json) {
        Write-Output $jsonText
    } else {
        Write-Output "Saved WPS Writer UI profile '$Profile' to $resolvedOutput ($(@($bars).Count) command bars, $(@($fieldCommands).Count) field commands)."
    }
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
