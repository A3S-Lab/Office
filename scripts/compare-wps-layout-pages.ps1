param(
    [Parameter(Mandatory = $true)]
    [string]$A3sPath,

    [Parameter(Mandatory = $true)]
    [string]$WpsPath,

    [Parameter(Mandatory = $true)]
    [string]$ReportPath,

    [ValidateRange(0, 1)]
    [double]$MaximumDifferentPixelRatio = 0.02,

    [ValidateRange(0, 255)]
    [double]$MaximumMeanAbsoluteError = 1.0,

    [ValidateRange(0, 10)]
    [int]$MaximumLandmarkDelta = 1
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function Get-ColorBounds {
    param(
        [Drawing.Bitmap]$Bitmap,
        [Drawing.Color]$Target,
        [int]$Tolerance,
        [int]$Left,
        [int]$Top,
        [int]$Right,
        [int]$Bottom,
        [switch]$Dark
    )

    $minimumX = [int]::MaxValue
    $minimumY = [int]::MaxValue
    $maximumX = -1
    $maximumY = -1
    $count = 0
    for ($y = $Top; $y -le $Bottom; $y += 1) {
        for ($x = $Left; $x -le $Right; $x += 1) {
            $color = $Bitmap.GetPixel($x, $y)
            $matches = if ($Dark) {
                $color.R -lt 100 -and $color.G -lt 100 -and $color.B -lt 100
            } else {
                [Math]::Abs($color.R - $Target.R) -le $Tolerance -and
                [Math]::Abs($color.G - $Target.G) -le $Tolerance -and
                [Math]::Abs($color.B - $Target.B) -le $Tolerance
            }
            if (-not $matches) { continue }
            $minimumX = [Math]::Min($minimumX, $x)
            $minimumY = [Math]::Min($minimumY, $y)
            $maximumX = [Math]::Max($maximumX, $x)
            $maximumY = [Math]::Max($maximumY, $y)
            $count += 1
        }
    }
    if ($count -eq 0) { throw 'A required visual landmark was not found.' }
    return [pscustomobject]@{
        x = $minimumX
        y = $minimumY
        right = $maximumX
        bottom = $maximumY
        pixels = $count
    }
}

function Get-DenseColorBand {
    param(
        [Drawing.Bitmap]$Bitmap,
        [string]$HexColor
    )

    $target = [Drawing.ColorTranslator]::FromHtml("#$HexColor")
    $denseRows = [Collections.Generic.List[int]]::new()
    for ($y = 140; $y -le 310; $y += 1) {
        $count = 0
        for ($x = 180; $x -le 600; $x += 1) {
            $color = $Bitmap.GetPixel($x, $y)
            if (
                [Math]::Abs($color.R - $target.R) -le 3 -and
                [Math]::Abs($color.G - $target.G) -le 3 -and
                [Math]::Abs($color.B - $target.B) -le 3
            ) {
                $count += 1
            }
        }
        if ($count -gt 200) { $denseRows.Add($y) }
    }
    if ($denseRows.Count -eq 0) {
        throw "The #$HexColor table band was not found."
    }
    $groups = [Collections.Generic.List[object]]::new()
    $start = $denseRows[0]
    $end = $start
    foreach ($row in $denseRows | Select-Object -Skip 1) {
        if ($row -eq $end + 1) {
            $end = $row
            continue
        }
        $groups.Add([pscustomobject]@{ top = $start; bottom = $end })
        $start = $row
        $end = $row
    }
    $groups.Add([pscustomobject]@{ top = $start; bottom = $end })
    return $groups |
        Sort-Object { $_.bottom - $_.top } -Descending |
        Select-Object -First 1
}

function Get-Landmarks {
    param([Drawing.Bitmap]$Bitmap)

    return [ordered]@{
        title = Get-ColorBounds -Bitmap $Bitmap `
            -Target ([Drawing.ColorTranslator]::FromHtml('#2E74B5')) `
            -Tolerance 8 -Left 80 -Top 80 -Right 350 -Bottom 140
        body = Get-ColorBounds -Bitmap $Bitmap -Target ([Drawing.Color]::Black) `
            -Tolerance 0 -Left 80 -Top 130 -Right 570 -Bottom 175 -Dark
        header = Get-DenseColorBand -Bitmap $Bitmap -HexColor '4472C4'
        bodyRow = Get-DenseColorBand -Bitmap $Bitmap -HexColor 'FCE4D6'
        totalRow = Get-DenseColorBand -Bitmap $Bitmap -HexColor 'E7E6E6'
        ending = Get-ColorBounds -Bitmap $Bitmap -Target ([Drawing.Color]::Black) `
            -Tolerance 0 -Left 80 -Top 280 -Right 250 -Bottom 320 -Dark
    }
}

function Get-MaximumLandmarkDelta {
    param(
        [Collections.IDictionary]$A3s,
        [Collections.IDictionary]$Wps
    )

    $maximum = 0
    foreach ($name in $A3s.Keys) {
        $properties = if ($name -in @('header', 'bodyRow', 'totalRow')) {
            @('top', 'bottom')
        } else {
            @('x', 'y')
        }
        foreach ($property in $properties) {
            $leftProperty = $A3s[$name].PSObject.Properties[$property]
            $rightProperty = $Wps[$name].PSObject.Properties[$property]
            $leftValue = if ($null -ne $leftProperty) { $leftProperty.Value } else { $null }
            $rightValue = if ($null -ne $rightProperty) { $rightProperty.Value } else { $null }
            if ($null -eq $leftValue -or $null -eq $rightValue) { continue }
            $maximum = [Math]::Max(
                $maximum,
                [Math]::Abs([double]$leftValue - [double]$rightValue)
            )
        }
    }
    return $maximum
}

$resolvedA3s = (Resolve-Path -LiteralPath $A3sPath).Path
$resolvedWps = (Resolve-Path -LiteralPath $WpsPath).Path
$resolvedReport = [IO.Path]::GetFullPath($ReportPath)
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($resolvedReport)) | Out-Null

$a3s = [Drawing.Bitmap]::new($resolvedA3s)
$wps = [Drawing.Bitmap]::new($resolvedWps)
try {
    if ($a3s.Width -ne $wps.Width -or $a3s.Height -ne $wps.Height) {
        throw "Page images differ in size: A3S $($a3s.Width)x$($a3s.Height), WPS $($wps.Width)x$($wps.Height)."
    }
    $differentPixels = 0
    [long]$absoluteError = 0
    for ($y = 0; $y -lt $a3s.Height; $y += 1) {
        for ($x = 0; $x -lt $a3s.Width; $x += 1) {
            $left = $a3s.GetPixel($x, $y)
            $right = $wps.GetPixel($x, $y)
            $red = [Math]::Abs($left.R - $right.R)
            $green = [Math]::Abs($left.G - $right.G)
            $blue = [Math]::Abs($left.B - $right.B)
            if ([Math]::Max($red, [Math]::Max($green, $blue)) -gt 16) {
                $differentPixels += 1
            }
            $absoluteError += $red + $green + $blue
        }
    }
    $pixelCount = $a3s.Width * $a3s.Height
    $differentPixelRatio = $differentPixels / $pixelCount
    $meanAbsoluteError = $absoluteError / ($pixelCount * 3)
    $a3sLandmarks = Get-Landmarks -Bitmap $a3s
    $wpsLandmarks = Get-Landmarks -Bitmap $wps
    $landmarkDelta = Get-MaximumLandmarkDelta -A3s $a3sLandmarks -Wps $wpsLandmarks
    $result = [ordered]@{
        a3s = $resolvedA3s
        wps = $resolvedWps
        width = $a3s.Width
        height = $a3s.Height
        differentPixelRatio = [Math]::Round($differentPixelRatio, 6)
        maximumDifferentPixelRatio = $MaximumDifferentPixelRatio
        meanAbsoluteError = [Math]::Round($meanAbsoluteError, 4)
        maximumMeanAbsoluteError = $MaximumMeanAbsoluteError
        maximumLandmarkDelta = $landmarkDelta
        allowedLandmarkDelta = $MaximumLandmarkDelta
        landmarks = [ordered]@{ a3s = $a3sLandmarks; wps = $wpsLandmarks }
        passed = (
            $differentPixelRatio -le $MaximumDifferentPixelRatio -and
            $meanAbsoluteError -le $MaximumMeanAbsoluteError -and
            $landmarkDelta -le $MaximumLandmarkDelta
        )
    }
    $json = $result | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText($resolvedReport, "$json`n")
    $json
    if (-not $result.passed) {
        throw 'A3S Office does not meet the WPS page-layout parity thresholds.'
    }
} finally {
    $a3s.Dispose()
    $wps.Dispose()
}
