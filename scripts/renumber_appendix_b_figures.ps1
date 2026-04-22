param(
  [Parameter(Mandatory = $true)][string]$InputDocx,
  [Parameter(Mandatory = $true)][string]$OutputDocx,
  [Parameter(Mandatory = $false)][string]$AppendixLetter
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

if (-not $AppendixLetter) {
  # Default: Ukrainian appendix letter "Б" (U+0411).
  $AppendixLetter = [string][char]0x0411
}

$wordAppendix =
  -join @(
    [char]0x0414, # U+0414
    [char]0x043E, # U+043E
    [char]0x0434, # U+0434
    [char]0x0430, # U+0430
    [char]0x0442, # U+0442
    [char]0x043E, # U+043E
    [char]0x043A  # U+043A
  )

$wordRis =
  -join @(
    [char]0x0420, # U+0420
    [char]0x0438, # U+0438
    [char]0x0441  # U+0441
  )

$wordRisDot = $wordRis + "."
$wordRisunok = $wordRis + (-join @([char]0x0443, [char]0x043D, [char]0x043E, [char]0x043A))

function New-DirectoryClean {
  param([Parameter(Mandatory = $true)][string]$Path)
  if (Test-Path -LiteralPath $Path) {
    Remove-Item -LiteralPath $Path -Recurse -Force
  }
  New-Item -ItemType Directory -Path $Path | Out-Null
}

function Expand-Docx {
  param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [Parameter(Mandatory = $true)][string]$DestinationDir
  )
  New-DirectoryClean -Path $DestinationDir
  [System.IO.Compression.ZipFile]::ExtractToDirectory((Resolve-Path -LiteralPath $DocxPath).Path, $DestinationDir)
}

function Compress-Docx {
  param(
    [Parameter(Mandatory = $true)][string]$SourceDir,
    [Parameter(Mandatory = $true)][string]$OutDocxPath
  )
  $outFull = [System.IO.Path]::GetFullPath($OutDocxPath)
  if (Test-Path -LiteralPath $outFull) { Remove-Item -LiteralPath $outFull -Force }
  [System.IO.Compression.ZipFile]::CreateFromDirectory((Resolve-Path -LiteralPath $SourceDir).Path, $outFull)
}

function Read-XmlUtf8 {
  param([Parameter(Mandatory = $true)][string]$Path)
  $utf8 = New-Object System.Text.UTF8Encoding($false)
  return [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Path).Path, $utf8)
}

function Load-XmlDocument {
  param([Parameter(Mandatory = $true)][string]$Path)
  $doc = New-Object System.Xml.XmlDocument
  $doc.PreserveWhitespace = $true
  $doc.LoadXml((Read-XmlUtf8 -Path $Path))
  return $doc
}

function Save-XmlDocumentUtf8NoBom {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Xml,
    [Parameter(Mandatory = $true)][string]$Path
  )
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  $settings = New-Object System.Xml.XmlWriterSettings
  $settings.Encoding = $utf8NoBom
  $settings.Indent = $false
  $settings.NewLineHandling = [System.Xml.NewLineHandling]::None

  $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
  $Xml.Save($writer)
  $writer.Flush()
  $writer.Close()
}

function New-WordNamespaceManager {
  param([Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Xml)
  $nsm = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
  [void]$nsm.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
  Write-Output -NoEnumerate $nsm
}

function Get-ParagraphText {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
  )
  $parts = @()
  foreach ($t in $Paragraph.SelectNodes(".//w:t", $Nsm)) {
    $parts += $t.InnerText
  }
  return (($parts -join "") -replace "\s+", " ").Trim()
}

function Set-ParagraphTextPlain {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][string]$Text,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
  )
  $tNodes = @($Paragraph.SelectNodes(".//w:t", $Nsm))
  if ($tNodes.Count -eq 0) { return $false }
  $first = $tNodes[0]
  $first.InnerText = $Text
  for ($k = $tNodes.Count - 1; $k -ge 1; $k--) {
    [void]$tNodes[$k].ParentNode.RemoveChild($tNodes[$k])
  }
  return $true
}

function Ensure-ParentDir {
  param([Parameter(Mandatory = $true)][string]$Path)
  $parent = Split-Path -Parent $Path
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent | Out-Null
  }
}

Ensure-ParentDir -Path $OutputDocx

$outParentDir = Split-Path -Parent $OutputDocx
if (-not $outParentDir) {
  $outParentDir = (Get-Location).Path
} else {
  $outParentDir = (Resolve-Path -LiteralPath $outParentDir).Path
}

$workRoot = Join-Path $outParentDir "tmp\\work"
$workDir = Join-Path $workRoot ("docx_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

try {
  Expand-Docx -DocxPath $InputDocx -DestinationDir $workDir

  $docXmlPath = Join-Path $workDir "word\\document.xml"
  $xml = Load-XmlDocument -Path $docXmlPath
  $nsm = New-WordNamespaceManager -Xml $xml

  $paras = $xml.SelectNodes("//w:body/w:p", $nsm)

  $idxAppendix = $null
  for ($i = 0; $i -lt $paras.Count; $i++) {
    $t = Get-ParagraphText -Paragraph $paras[$i] -Nsm $nsm
    if (-not $t) { continue }
    if (-not $t.StartsWith($wordAppendix, [System.StringComparison]::Ordinal)) { continue }
    $after = $t.Substring($wordAppendix.Length).TrimStart()
    if ($after.Length -lt 2) { continue }
    if ($after.Substring(0, 1) -ne $AppendixLetter) { continue }
    $next = $after.Substring(1, 1)
    if ($next -match "[\p{L}\p{Nd}_]") { continue } # skip "B1" style headings
    $idxAppendix = $i
    break
  }

  if ($idxAppendix -eq $null) {
    throw ("Appendix '{0}' heading not found (expected a paragraph starting with '{1} {0}')." -f $AppendixLetter, $wordAppendix)
  }

  $idxEnd = $paras.Count
  for ($i = $idxAppendix + 1; $i -lt $paras.Count; $i++) {
    $t = Get-ParagraphText -Paragraph $paras[$i] -Nsm $nsm
    if (-not $t) { continue }
    if (-not $t.StartsWith($wordAppendix, [System.StringComparison]::Ordinal)) { continue }
    $after = $t.Substring($wordAppendix.Length).TrimStart()
    if ($after.Length -lt 2) { continue }
    $letter = $after.Substring(0, 1)
    if ($letter -notmatch "^\p{L}$") { continue }
    $next = $after.Substring(1, 1)
    if ($next -match "[\p{L}\p{Nd}_]") { continue } # skip "B1" style headings
    if ($letter -ne $AppendixLetter) {
      $idxEnd = $i
      break
    }
  }

  $captionPattern =
    '^(?:' +
    [Regex]::Escape($wordRisDot) +
    '|' +
    [Regex]::Escape($wordRisunok) +
    ')\s*(?<num>[^\s]+)(?<rest>.*)$'
  $captionRegex = [regex]$captionPattern

  $counter = 0
  for ($i = $idxAppendix; $i -lt $idxEnd; $i++) {
    $t = Get-ParagraphText -Paragraph $paras[$i] -Nsm $nsm
    if (-not $t) { continue }

    $m = $captionRegex.Match($t)
    if (-not $m.Success) { continue }

    $counter++
    $rest = $m.Groups["rest"].Value
    if ($rest) {
      $first = $rest[0]
      $ok =
        [char]::IsWhiteSpace($first) -or
        ($first -eq '-') -or
        ($first -eq [char]0x2014)
      if (-not $ok) { $rest = " " + $rest }
    }
    $newText = ("{0} {1}.{2}{3}" -f $wordRisDot, $AppendixLetter, $counter, $rest)

    [void](Set-ParagraphTextPlain -Paragraph $paras[$i] -Text $newText -Nsm $nsm)
  }

  if ($counter -eq 0) {
    throw ("No figure captions found in Appendix '{0}'." -f $AppendixLetter)
  }

  Save-XmlDocumentUtf8NoBom -Xml $xml -Path $docXmlPath
  Compress-Docx -SourceDir $workDir -OutDocxPath $OutputDocx

  Write-Host ("Renumbered {0} figure captions in Appendix {1} -> {2}" -f $counter, $AppendixLetter, (Split-Path -Leaf $OutputDocx))
} finally {
  if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
  }
}
