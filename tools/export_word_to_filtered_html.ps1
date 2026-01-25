param(
  [Parameter(Mandatory = $true)]
  [string]$DocPath,

  [Parameter(Mandatory = $true)]
  [string]$OutDir
)

$ErrorActionPreference = 'Stop'

$resolvedDocPath = (Resolve-Path -LiteralPath $DocPath).Path
$resolvedOutDir = (Resolve-Path -LiteralPath $OutDir -ErrorAction SilentlyContinue)
if (-not $resolvedOutDir) {
  New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
  $resolvedOutDir = (Resolve-Path -LiteralPath $OutDir).Path
} else {
  $resolvedOutDir = $resolvedOutDir.Path
}

$htmlPath = Join-Path $resolvedOutDir 'word.htm'
$wdFormatFilteredHTML = 10

$word = $null
$doc = $null

try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.AutomationSecurity = 3

  Write-Output 'STEP:OPEN'
  $doc = $word.Documents.Open(
    $resolvedDocPath,
    $false,
    $true,
    $false,
    '',
    '',
    $false,
    '',
    '',
    0,
    0,
    $false,
    $true,
    0,
    $true
  )
  Write-Output 'STEP:SAVE'
  $doc.SaveAs2($htmlPath, $wdFormatFilteredHTML)
  Write-Output 'STEP:DONE'
}
finally {
  if ($doc) {
    try { $doc.Close() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
  }
  if ($word) {
    try { $word.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
}

if (-not (Test-Path -LiteralPath $htmlPath)) {
  throw "Export failed: $htmlPath not created"
}

Write-Output $htmlPath

