param(
  [Parameter(Mandatory=$true)]
  [string]$SourceFolder,
  [string]$DestFolder = ""
)

if ([string]::IsNullOrWhiteSpace($DestFolder)) { $DestFolder = $SourceFolder }

$wdFormatXMLDocument = 16  # docx
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
  $docs = Get-ChildItem -Path $SourceFolder -Recurse -File |
          Where-Object { $_.Extension -ieq ".doc" -and $_.Name -notmatch "^\~\$" }

  foreach ($f in $docs) {
    $relative = $f.FullName.Substring($SourceFolder.Length).TrimStart("\","/")
    $outPath  = Join-Path $DestFolder ([System.IO.Path]::ChangeExtension($relative, ".docx"))

    $outDir = Split-Path $outPath -Parent
    if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

    Write-Host "Converto: $($f.FullName) -> $outPath"

    # Apri in sola lettura
    $doc = $word.Documents.Open($f.FullName, $false, $true)

    try {
      # IMPORTANTISSIMO: passare stringhe "pure"
      $doc.SaveAs2([string]$outPath, [int]$wdFormatXMLDocument)
    }
    finally {
      $doc.Close($false) | Out-Null
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
  }
}
finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

Write-Host "Fatto."

