param(
    [Parameter(Mandatory = $true)]
    [string]$A3sPath,

    [Parameter(Mandatory = $true)]
    [string]$WpsPath,

    [Parameter(Mandatory = $true)]
    [string]$ReportPath,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 200)]
    [int]$ExpectedBandCount,

    [ValidateRange(0, 10)]
    [int]$MaximumBandTopDelta = 3,

    [ValidateRange(0, 10)]
    [int]$MaximumBandAdvanceDelta = 4
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function Get-TextBands {
    param(
        [Drawing.Bitmap]$Bitmap,
        [int]$ExpectedCount
    )

    $minimumDarkChannels = 140
    $minimumRowPixels = 3
    $minimumPeakPixels = 8
    $rows = [Collections.Generic.List[object]]::new()
    $darkPixels = 0
    for ($y = 60; $y -lt $Bitmap.Height - 40; $y += 1) {
        $count = 0
        for ($x = 70; $x -lt $Bitmap.Width - 40; $x += 1) {
            $color = $Bitmap.GetPixel($x, $y)
            if (
                $color.R -lt $minimumDarkChannels -and
                $color.G -lt $minimumDarkChannels -and
                $color.B -lt $minimumDarkChannels
            ) {
                $count += 1
            }
        }
        $darkPixels += $count
        if ($count -ge $minimumRowPixels) {
            $rows.Add([pscustomobject]@{ y = $y; pixels = $count })
        }
    }

    $selectedGap = 3
    $bands = @(Merge-TextRows `
        -Rows $rows `
        -MaximumBlankRowGap $selectedGap `
        -MinimumPeakPixels $minimumPeakPixels)
    if ($bands.Count -ne $ExpectedCount) {
        foreach ($candidateGap in @(2, 4, 1, 5, 0, 6)) {
            $candidate = @(Merge-TextRows `
                -Rows $rows `
                -MaximumBlankRowGap $candidateGap `
                -MinimumPeakPixels $minimumPeakPixels)
            if ($candidate.Count -eq $ExpectedCount) {
                $selectedGap = $candidateGap
                $bands = $candidate
                break
            }
        }
    }

    return [pscustomobject]@{
        bands = $bands
        maximumBlankRowGap = $selectedGap
        darkPixels = $darkPixels
    }
}

function Merge-TextRows {
    param(
        [object[]]$Rows,
        [int]$MaximumBlankRowGap,
        [int]$MinimumPeakPixels
    )

    $bands = [Collections.Generic.List[object]]::new()
    if ($Rows.Count -eq 0) { return $bands }
    $start = $Rows[0].y
    $end = $start
    $peak = $Rows[0].pixels
    foreach ($row in $Rows | Select-Object -Skip 1) {
        if ($row.y -le $end + $MaximumBlankRowGap + 1) {
            $end = $row.y
            $peak = [Math]::Max($peak, $row.pixels)
            continue
        }
        if ($peak -ge $MinimumPeakPixels) {
            $bands.Add([pscustomobject]@{
                top = $start
                bottom = $end
                peakPixels = $peak
            })
        }
        $start = $row.y
        $end = $row.y
        $peak = $row.pixels
    }
    if ($peak -ge $MinimumPeakPixels) {
        $bands.Add([pscustomobject]@{
            top = $start
            bottom = $end
            peakPixels = $peak
        })
    }
    return $bands
}

$resolvedA3s = (Resolve-Path -LiteralPath $A3sPath).Path
$resolvedWps = (Resolve-Path -LiteralPath $WpsPath).Path
$resolvedReport = [IO.Path]::GetFullPath($ReportPath)
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($resolvedReport)) |
    Out-Null

$a3s = [Drawing.Bitmap]::new($resolvedA3s)
$wps = [Drawing.Bitmap]::new($resolvedWps)
try {
    if ($a3s.Width -ne $wps.Width -or $a3s.Height -ne $wps.Height) {
        throw "Page images differ in size: A3S $($a3s.Width)x$($a3s.Height), WPS $($wps.Width)x$($wps.Height)."
    }
    $a3sResult = Get-TextBands `
        -Bitmap $a3s `
        -ExpectedCount $ExpectedBandCount
    $wpsResult = Get-TextBands `
        -Bitmap $wps `
        -ExpectedCount $ExpectedBandCount
    $sameBandCount =
        $a3sResult.bands.Count -eq $ExpectedBandCount -and
        $wpsResult.bands.Count -eq $ExpectedBandCount
    $maximumTopDelta = 0
    $maximumBottomDelta = 0
    $maximumAdvanceDelta = 0
    if ($sameBandCount) {
        for ($index = 0; $index -lt $ExpectedBandCount; $index += 1) {
            $a3sBand = $a3sResult.bands[$index]
            $wpsBand = $wpsResult.bands[$index]
            $maximumTopDelta = [Math]::Max(
                $maximumTopDelta,
                [Math]::Abs($a3sBand.top - $wpsBand.top)
            )
            $maximumBottomDelta = [Math]::Max(
                $maximumBottomDelta,
                [Math]::Abs($a3sBand.bottom - $wpsBand.bottom)
            )
            if ($index -eq 0) { continue }
            $a3sAdvance = $a3sBand.top - $a3sResult.bands[$index - 1].top
            $wpsAdvance = $wpsBand.top - $wpsResult.bands[$index - 1].top
            $maximumAdvanceDelta = [Math]::Max(
                $maximumAdvanceDelta,
                [Math]::Abs($a3sAdvance - $wpsAdvance)
            )
        }
    }
    $passed =
        $sameBandCount -and
        $maximumTopDelta -le $MaximumBandTopDelta -and
        $maximumAdvanceDelta -le $MaximumBandAdvanceDelta
    $result = [ordered]@{
        a3s = $resolvedA3s
        wps = $resolvedWps
        width = $a3s.Width
        height = $a3s.Height
        expectedBandCount = $ExpectedBandCount
        a3sBandCount = $a3sResult.bands.Count
        wpsBandCount = $wpsResult.bands.Count
        maximumBlankRowGap = [ordered]@{
            a3s = $a3sResult.maximumBlankRowGap
            wps = $wpsResult.maximumBlankRowGap
        }
        maximumBandTopDelta = $maximumTopDelta
        allowedBandTopDelta = $MaximumBandTopDelta
        maximumBandBottomDelta = $maximumBottomDelta
        maximumBandAdvanceDelta = $maximumAdvanceDelta
        allowedBandAdvanceDelta = $MaximumBandAdvanceDelta
        darkPixels = [ordered]@{
            a3s = $a3sResult.darkPixels
            wps = $wpsResult.darkPixels
        }
        bands = [ordered]@{
            a3s = $a3sResult.bands
            wps = $wpsResult.bands
        }
        passed = $passed
    }
    $json = $result | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText($resolvedReport, "$json`n")
    $json
    if (-not $passed) {
        throw 'A3S Office does not meet the WPS text-layout parity thresholds.'
    }
} finally {
    $a3s.Dispose()
    $wps.Dispose()
}
