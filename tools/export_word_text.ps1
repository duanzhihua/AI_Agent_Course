param(
  [Parameter(Mandatory = $true)]
  [string]$DocPath,

  [Parameter(Mandatory = $true)]
  [string]$OutPath
)

$ErrorActionPreference = 'Stop'

$resolvedDocPath = (Resolve-Path -LiteralPath $DocPath).Path
$resolvedOutPath = $OutPath
$resolvedOutDir = Split-Path -Parent $resolvedOutPath
if (-not (Test-Path -LiteralPath $resolvedOutDir)) {
  New-Item -ItemType Directory -Force -Path $resolvedOutDir | Out-Null
}

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

  $text = $doc.Content.Text
  [IO.File]::WriteAllText($resolvedOutPath, $text, [Text.UTF8Encoding]::new($false))
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

if (-not (Test-Path -LiteralPath $resolvedOutPath)) {
  throw "Export failed: $resolvedOutPath not created"
}

Write-Output $resolvedOutPath

