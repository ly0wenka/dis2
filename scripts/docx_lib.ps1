param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

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
  $out = (Resolve-Path -LiteralPath (Split-Path -Parent $OutDocxPath)).Path + "\" + (Split-Path -Leaf $OutDocxPath)
  if (Test-Path -LiteralPath $out) {
    Remove-Item -LiteralPath $out -Force
  }
  [System.IO.Compression.ZipFile]::CreateFromDirectory((Resolve-Path -LiteralPath $SourceDir).Path, $out)
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
  [void]$nsm.AddNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math")
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

function Paragraph-HasMath {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
  )
  return ($Paragraph.SelectNodes(".//m:oMath | .//m:oMathPara", $Nsm).Count -gt 0)
}

function Paragraph-IsHeadingText {
  param([AllowEmptyString()][string[]]$Text)
  $s = ($Text -join "")
  return ($s -match "^\d+\.\d+(\.\d+)?\s")
}

function Paragraph-IsCaptionText {
  param([AllowEmptyString()][string[]]$Text)
  $s = ($Text -join "")
  return ($s -match "^(Рис\.|Табл\.|Таблиця|Рисунок)\s")
}

function Paragraph-IsOnlyEquationNumberText {
  param([AllowEmptyString()][string[]]$Text)
  $s = ($Text -join "")
  return ($s -match "^\(\d+\.\d+\)\s*$")
}

function Normalize-BracketCitationsInText {
  param([AllowEmptyString()][string[]]$Text)
  $s = ($Text -join "")
  if ($s.IndexOf("[") -lt 0 -or $s.IndexOf("]") -lt 0 -or $s.IndexOf(";") -lt 0) {
    return $s
  }
  $sb = New-Object System.Text.StringBuilder
  $inBracket = $false
  for ($i = 0; $i -lt $s.Length; $i++) {
    $ch = $s[$i]
    if ($ch -eq "[") { $inBracket = $true }
    elseif ($ch -eq "]") { $inBracket = $false }

    if ($inBracket -and $ch -eq ";") {
      [void]$sb.Append(",")
      continue
    }
    [void]$sb.Append($ch)
  }
  return $sb.ToString()
}

function Remove-RefTokensInText {
  param([AllowEmptyString()][string[]]$Text)
  $s = ($Text -join "")
  if ($s -notmatch "REF\s+_Ref") { return $s }
  $s = $s -replace "REF\s+_Ref[0-9A-Za-z_]+", ""
  $s = $s -replace "REF\s+_Ref", ""
  $s = $s -replace "\s{2,}", " "
  return $s.Trim()
}
