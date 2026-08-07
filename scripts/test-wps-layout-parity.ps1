param(
    [string]$BaseUrl = 'http://127.0.0.1:4175',
    [string]$ChromiumPath = $env:A3S_OFFICE_VISUAL_CHROMIUM_EXECUTABLE,
    [string]$OutputDirectory = '.a3s-test/wps-layout-parity'
)

$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$output = [IO.Path]::GetFullPath((Join-Path $root $OutputDirectory))

function Invoke-WpsParityCapture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Fixture,

        [int]$ExpectedBandCount = 0
    )

    $pdf = Join-Path $output "$Name.pdf"
    $a3sPng = Join-Path $output "a3s-$Name.png"
    $wpsPng = Join-Path $output "wps-$Name.png"
    $layoutJson = Join-Path $output "$Name-browser-layout.json"
    $comparisonJson = Join-Path $output "$Name-comparison.json"

    & (Join-Path $PSScriptRoot 'export-wps-reference.ps1') `
        -InputPath $Fixture `
        -OutputPath $pdf `
        -ExpectedPageCount 1 |
        Tee-Object -FilePath (Join-Path $output "$Name-wps-export.json")

    $captureArguments = @(
        'scripts/capture-wps-layout-pages.mjs',
        '--base-url', $BaseUrl,
        '--docx', $Fixture,
        '--pdf', $pdf,
        '--a3s-output', $a3sPng,
        '--wps-output', $wpsPng,
        '--layout-output', $layoutJson
    )
    if ($ChromiumPath) {
        $captureArguments += @('--chromium', $ChromiumPath)
    }
    & node @captureArguments
    if ($LASTEXITCODE -ne 0) { throw "Unable to capture $Name parity pages." }

    if ($ExpectedBandCount -gt 0) {
        & (Join-Path $PSScriptRoot 'compare-wps-text-layout.ps1') `
            -A3sPath $a3sPng `
            -WpsPath $wpsPng `
            -ReportPath $comparisonJson `
            -ExpectedBandCount $ExpectedBandCount
        return
    }
    & (Join-Path $PSScriptRoot 'compare-wps-layout-pages.ps1') `
        -A3sPath $a3sPng `
        -WpsPath $wpsPng `
        -ReportPath $comparisonJson
}

[IO.Directory]::CreateDirectory($output) | Out-Null
Push-Location $root
try {
    & npx --yes bun@1.3.14 scripts/create-e2e-fixtures.ts
    if ($LASTEXITCODE -ne 0) { throw 'Unable to generate the WPS layout fixture.' }

    Invoke-WpsParityCapture `
        -Name 'word-wps-layout' `
        -Fixture (Join-Path $root '.a3s-test/fixtures/word-wps-layout.docx')
    Invoke-WpsParityCapture `
        -Name 'word-wps-font-matrix' `
        -Fixture (Join-Path $root '.a3s-test/fixtures/word-wps-font-matrix.docx') `
        -ExpectedBandCount 30
    Invoke-WpsParityCapture `
        -Name 'word-wps-cjk-font-matrix' `
        -Fixture (Join-Path $root '.a3s-test/fixtures/word-wps-cjk-font-matrix.docx') `
        -ExpectedBandCount 36
    Invoke-WpsParityCapture `
        -Name 'word-wps-grid-matrix' `
        -Fixture (Join-Path $root '.a3s-test/fixtures/word-wps-grid-matrix.docx') `
        -ExpectedBandCount 18
    Invoke-WpsParityCapture `
        -Name 'word-wps-script-matrix' `
        -Fixture (Join-Path $root '.a3s-test/fixtures/word-wps-script-matrix.docx') `
        -ExpectedBandCount 30
} finally {
    Pop-Location
}
