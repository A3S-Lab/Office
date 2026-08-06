param(
    [string]$BaseUrl = 'http://127.0.0.1:4175',
    [string]$ChromiumPath = $env:A3S_OFFICE_VISUAL_CHROMIUM_EXECUTABLE,
    [string]$OutputDirectory = '.a3s-test/wps-layout-parity'
)

$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$output = [IO.Path]::GetFullPath((Join-Path $root $OutputDirectory))
$fixture = Join-Path $root '.a3s-test/fixtures/word-wps-layout.docx'
$pdf = Join-Path $output 'word-wps-layout.pdf'
$a3sPng = Join-Path $output 'a3s-word-wps-layout.png'
$wpsPng = Join-Path $output 'wps-word-wps-layout.png'
$layoutJson = Join-Path $output 'browser-layout.json'
$comparisonJson = Join-Path $output 'comparison.json'

[IO.Directory]::CreateDirectory($output) | Out-Null
Push-Location $root
try {
    & npx --yes bun@1.3.14 scripts/create-e2e-fixtures.ts
    if ($LASTEXITCODE -ne 0) { throw 'Unable to generate the WPS layout fixture.' }

    & (Join-Path $PSScriptRoot 'export-wps-reference.ps1') `
        -InputPath $fixture `
        -OutputPath $pdf `
        -ExpectedPageCount 1 |
        Tee-Object -FilePath (Join-Path $output 'wps-export.json')

    $captureArguments = @(
        'scripts/capture-wps-layout-pages.mjs',
        '--base-url', $BaseUrl,
        '--docx', $fixture,
        '--pdf', $pdf,
        '--a3s-output', $a3sPng,
        '--wps-output', $wpsPng,
        '--layout-output', $layoutJson
    )
    if ($ChromiumPath) {
        $captureArguments += @('--chromium', $ChromiumPath)
    }
    & node @captureArguments
    if ($LASTEXITCODE -ne 0) { throw 'Unable to capture the parity pages.' }

    & (Join-Path $PSScriptRoot 'compare-wps-layout-pages.ps1') `
        -A3sPath $a3sPng `
        -WpsPath $wpsPng `
        -ReportPath $comparisonJson
} finally {
    Pop-Location
}
