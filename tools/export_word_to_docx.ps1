param(
  [Parameter(Mandatory = $true)]
  [string]$DocPath,

  [Parameter(Mandatory = $true)]
  [string]$DocxPath
)

$ErrorActionPreference = 'Stop'

$resolvedDocPath = (Resolve-Path -LiteralPath $DocPath).Path
$resolvedDocxPath = $DocxPath
$resolvedDocxDir = Split-Path -Parent $resolvedDocxPath
if (-not (Test-Path -LiteralPath $resolvedDocxDir)) {
  New-Item -ItemType Directory -Force -Path $resolvedDocxDir | Out-Null
}

$wdFormatXMLDocument = 12

$word = $null
$doc = $null

try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.AutomationSecurity = 3

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

  $doc.SaveAs2($resolvedDocxPath, $wdFormatXMLDocument)
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

if (-not (Test-Path -LiteralPath $resolvedDocxPath)) {
  throw "Export failed: $resolvedDocxPath not created"
}

Write-Output $resolvedDocxPath

