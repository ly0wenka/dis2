param(
  [Parameter(Mandatory = $true)][string]$InputDocx,
  [Parameter(Mandatory = $true)][string]$OutputDocx,
  [Parameter(Mandatory = $true)][string]$ReportPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

. "$PSScriptRoot\\docx_lib.ps1"

function Ensure-ParentDir {
  param([Parameter(Mandatory = $true)][string]$Path)
  $parent = Split-Path -Parent $Path
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }
}

function Paragraph-HasDrawing {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
  )
  return ($Paragraph.SelectSingleNode(".//w:drawing | .//w:pict | .//w:object", $Nsm) -ne $null)
}

function Paragraph-GetStyle {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
  )
  $attr = $Paragraph.SelectSingleNode("./w:pPr/w:pStyle/@w:val", $Nsm)
  if ($attr) { return $attr.Value }
  return ""
}

function Paragraph-SetStyle {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm,
    [Parameter(Mandatory = $true)][string]$StyleVal
  )
  $pPr = $Paragraph.SelectSingleNode("./w:pPr", $Nsm)
  if (-not $pPr) {
    $pPr = $Paragraph.OwnerDocument.CreateElement("w", "pPr", $Nsm.LookupNamespace("w"))
    [void]$Paragraph.PrependChild($pPr)
  }
  $pStyle = $pPr.SelectSingleNode("./w:pStyle", $Nsm)
  if (-not $pStyle) {
    $pStyle = $Paragraph.OwnerDocument.CreateElement("w", "pStyle", $Nsm.LookupNamespace("w"))
    [void]$pPr.PrependChild($pStyle)
  }
  $pStyle.SetAttribute("w:val", $Nsm.LookupNamespace("w"), $StyleVal)
}

function Rewrite-TextOnlyParagraph {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm,
    [Parameter(Mandatory = $true)][string]$Text
  )
  # Keep pPr (style), replace all runs with a single clean run (inherits formatting from style).
  foreach ($r in @($Paragraph.SelectNodes("./w:r", $Nsm))) {
    [void]$Paragraph.RemoveChild($r)
  }
  $rNew = $Paragraph.OwnerDocument.CreateElement("w", "r", $Nsm.LookupNamespace("w"))
  $tNew = $Paragraph.OwnerDocument.CreateElement("w", "t", $Nsm.LookupNamespace("w"))
  $tNew.SetAttribute("xml:space", "preserve")
  $tNew.InnerText = $Text
  [void]$rNew.AppendChild($tNew)
  [void]$Paragraph.AppendChild($rNew)
}

function Insert-ParagraphAfterIndex {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Xml,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm,
    [Parameter(Mandatory = $true)][System.Xml.XmlNodeList]$Paras,
    [Parameter(Mandatory = $true)][int]$AfterIndex,
    [Parameter(Mandatory = $true)][string]$StyleVal,
    [Parameter(Mandatory = $true)][string]$Text
  )
  $pNew = $Xml.CreateElement("w", "p", $Nsm.LookupNamespace("w"))
  $pPr = $Xml.CreateElement("w", "pPr", $Nsm.LookupNamespace("w"))
  $pStyle = $Xml.CreateElement("w", "pStyle", $Nsm.LookupNamespace("w"))
  $pStyle.SetAttribute("w:val", $Nsm.LookupNamespace("w"), $StyleVal)
  [void]$pPr.AppendChild($pStyle)
  [void]$pNew.AppendChild($pPr)
  Rewrite-TextOnlyParagraph -Paragraph $pNew -Nsm $Nsm -Text $Text
  [void]$Paras[$AfterIndex].ParentNode.InsertAfter($pNew, $Paras[$AfterIndex])
}

function Normalize-BracketCitationsInParagraphRuns {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
  )
  $tNodes = @($Paragraph.SelectNodes(".//w:t", $Nsm))
  if ($tNodes.Count -eq 0) { return 0 }
  $inBracket = $false
  $replaced = 0
  foreach ($t in $tNodes) {
    $s = $t.InnerText
    if (-not $s) { continue }
    if ($s.IndexOf("[") -lt 0 -and $s.IndexOf("]") -lt 0 -and $s.IndexOf(";") -lt 0) { continue }
    $chars = $s.ToCharArray()
    for ($i = 0; $i -lt $chars.Length; $i++) {
      if ($chars[$i] -eq "[") { $inBracket = $true; continue }
      if ($chars[$i] -eq "]") { $inBracket = $false; continue }
      if ($inBracket -and $chars[$i] -eq ";") { $chars[$i] = ","; $replaced++ }
    }
    $newS = -join $chars
    if ($newS -ne $s) { $t.InnerText = $newS }
  }
  return $replaced
}

function Wrap-BareNumericListsInTextNode {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$TextNode,
    [Parameter(Mandatory = $true)][ref]$SkippedOut
  )
  $s = $TextNode.InnerText
  if (-not $s) { return 0 }
  if ($s -notmatch "\d{1,3}\s*;\s*\d{1,3}") { return 0 }
  $script:__wrapCount = 0
  $pattern = '(\b\d{1,3}(?:\s*;\s*\d{1,3}){1,})(?=\s*[\)\].,;:]|\s*$)'
  $new = [System.Text.RegularExpressions.Regex]::Replace($s, $pattern, {
    param($m)
    $list = $m.Groups[1].Value
    if ($list -match '[\\.-]') {
      $SkippedOut.Value += $list
      return $list
    }
    $items = ($list -split ';' | ForEach-Object { $_.Trim() }) -join ', '
    $script:__wrapCount++
    return '[' + $items + ']'
  })
  if ($new -ne $s) { $TextNode.InnerText = $new }
  return $script:__wrapCount
}

function Split-Sentences {
  param([Parameter(Mandatory = $true)][string]$Text)
  $t = ($Text -replace "\s{2,}", " ").Trim()
  if (-not $t) { return @() }
  $parts = [System.Text.RegularExpressions.Regex]::Split($t, "(?<=[\\.!?])\\s+")
  return @($parts | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Clean-BoilerplateFromText {
  param([Parameter(Mandatory = $true)][string]$Text)
  $rx = [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
  $t = $Text
  # Panel/detection phrases.
  $t = [regex]::Replace($t, "Ð”ÐµÑ‚ÐµÐºÑ†Ñ–Ñ\\s+Ð¾Ð±â€™Ñ”ÐºÑ‚Ñ–Ð²\\.?\\s*", "", $rx)
  $t = [regex]::Replace($t, "Ð’Ð¸ÑÐ²Ð»ÐµÐ½Ð¾\\s+\\d+\\s+Ð¾Ð±.?Ñ”ÐºÑ‚Ñ–Ð²\\s+Ð¼Ð¾Ð´ÐµÐ»Ð»ÑŽ\\s+DETR\\.?\\s*", "", $rx)
  $t = [regex]::Replace($t, "Ð£\\s+Ð½Ð¸Ð¶Ð½Ñ–Ð¹\\s+Ñ‡Ð°ÑÑ‚Ð¸Ð½Ñ–.*?\\.\\s*", "", $rx)
  $t = [regex]::Replace($t, "Ð›Ñ–Ð²Ð°\\s+Ð¿Ð°Ð½ÐµÐ»ÑŒ.*?\\.\\s*", "", $rx)
  $t = [regex]::Replace($t, "Ð¦ÐµÐ½Ñ‚Ñ€Ð°Ð»ÑŒÐ½Ð°\\s+Ð¿Ð°Ð½ÐµÐ»ÑŒ.*?\\.\\s*", "", $rx)
  $t = [regex]::Replace($t, "ÐŸÑ€Ð°Ð²Ð°\\s+Ð¿Ð°Ð½ÐµÐ»ÑŒ.*?\\.\\s*", "", $rx)
  $t = [regex]::Replace($t, "ÐÐ°\\s+Ð»Ñ–Ð²Ñ–Ð¹\\s+Ð¿Ð°Ð½ÐµÐ»Ñ–.*?\\.\\s*", "", $rx)
  $t = [regex]::Replace($t, "ÐÐ°\\s+Ñ†ÐµÐ½Ñ‚Ñ€Ð°Ð»ÑŒÐ½Ñ–Ð¹\\s+Ð¿Ð°Ð½ÐµÐ»Ñ–.*?\\.\\s*", "", $rx)
  $t = [regex]::Replace($t, "ÐÐ°\\s+Ð¿Ñ€Ð°Ð²Ñ–Ð¹\\s+Ð¿Ð°Ð½ÐµÐ»Ñ–.*?\\.\\s*", "", $rx)
  # Tidy.
  $t = ($t -replace "\s{2,}", " ").Trim()
  return $t
}

function Extract-AndStripCitations {
  param([Parameter(Mandatory = $true)][string]$Sentence, [Parameter(Mandatory = $true)][ref]$Cites)
  $s = $Sentence
  foreach ($m in [System.Text.RegularExpressions.Regex]::Matches($s, "\\[[^\\]]+\\]")) {
    $Cites.Value += $m.Value
  }
  $s = [regex]::Replace($s, "\\s*\\[[^\\]]+\\]\\s*", " ")
  $s = ($s -replace "\s{2,}", " ").Trim()
  return $s
}

function Merge-CitationBlocks {
  param([Parameter(Mandatory = $true)][string[]]$Blocks)
  if ($Blocks.Count -eq 0) { return "" }
  $items = New-Object 'System.Collections.Generic.List[string]'
  foreach ($b in $Blocks) {
    $inner = $b.Trim()
    if ($inner.StartsWith("[")) { $inner = $inner.Substring(1) }
    if ($inner.EndsWith("]")) { $inner = $inner.Substring(0, $inner.Length - 1) }
    foreach ($it in ($inner -split ",")) {
      $x = $it.Trim()
      if (-not $x) { continue }
      if (-not $items.Contains($x)) { [void]$items.Add($x) }
    }
  }
  if ($items.Count -eq 0) { return "" }
  return "[" + (($items.ToArray()) -join ", ") + "]"
}

function Find-BodyHeadingIndexByPrefix {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNodeList]$Paras,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm,
    [Parameter(Mandatory = $true)][string]$Prefix,
    [Parameter(Mandatory = $false)][int]$StartAt = 0,
    [Parameter(Mandatory = $false)][int]$Lookahead = 200
  )
  $cands = New-Object 'System.Collections.Generic.List[int]'
  for ($j = $StartAt; $j -lt $Paras.Count; $j++) {
    $txt = Get-ParagraphText -Paragraph $Paras[$j] -Nsm $Nsm
    if ($txt -and $txt.StartsWith($Prefix)) { [void]$cands.Add($j) }
  }
  if ($cands.Count -eq 0) { return $null }
  if ($cands.Count -eq 1) { return $cands[0] }
  foreach ($c in $cands) {
    $limit = [Math]::Min($Paras.Count, $c + $Lookahead)
    for ($k = $c + 1; $k -lt $limit; $k++) {
      $t = Get-ParagraphText -Paragraph $Paras[$k] -Nsm $Nsm
      if (-not $t) { continue }
      if ($t.StartsWith("4.5")) { break }
      if ($t -match "^(Ð Ð¸Ñ\\.|Ð Ð¸ÑÑƒÐ½Ð¾Ðº)\\s") { return $c }
    }
  }
  return $cands[$cands.Count - 1]
}

function Get-FigureInfoFromCaption {
  param([Parameter(Mandatory = $true)][string]$CaptionText)
  $m = [regex]::Match($CaptionText, "^Ð Ð¸Ñ\\.\\s*4\\.(\\d+)\\s*â€”\\s*(.+)$")
  if (-not $m.Success) { return $null }
  $figNo = [int]$m.Groups[1].Value
  $label = $m.Groups[2].Value.Trim()
  $prefix = $label
  $u = $label.IndexOf("_")
  if ($u -gt 0) { $prefix = $label.Substring(0, $u) }
  else {
    $sp = $label.IndexOf(" ")
    if ($sp -gt 0) { $prefix = $label.Substring(0, $sp) }
  }
  return [pscustomobject]@{ fig = $figNo; label = $label; prefix = $prefix }
}

function Remove-ParagraphAt {
  param([Parameter(Mandatory = $true)][System.Xml.XmlNodeList]$Paras, [Parameter(Mandatory = $true)][int]$Index)
  [void]$Paras[$Index].ParentNode.RemoveChild($Paras[$Index])
}

$workRoot = Join-Path (Split-Path -Parent $OutputDocx) "tmp\\work"
$workDir = Join-Path $workRoot ("docx_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

$stats = [ordered]@{
  input = (Resolve-Path -LiteralPath $InputDocx).Path
  output = (Resolve-Path -LiteralPath (Split-Path -Parent $OutputDocx)).Path + "\\" + (Split-Path -Leaf $OutputDocx)
  bracket_semicolons_replaced = 0
  bare_lists_wrapped = 0
  bare_lists_skipped = @()
  legends_inserted = 0
  legends_removed = 0
  panel_paras_deleted = 0
  detection_paras_deleted = 0
  rewritten_blocks = 0
  rewritten_paras = 0
  deleted_text_paras = 0
  style_body_used = @()
  style_caption_used = @()
}

try {
  # Copy input -> output (never edit input in place).
  Copy-Item -LiteralPath $InputDocx -Destination $OutputDocx -Force

  Expand-Docx -DocxPath $OutputDocx -DestinationDir $workDir
  $docXmlPath = Join-Path $workDir "word\\document.xml"
  $xml = Load-XmlDocument -Path $docXmlPath
  $nsm = New-WordNamespaceManager -Xml $xml

  $paras = $xml.SelectNodes("//w:body/w:p", $nsm)

  # Whole-document citation safety: bracket ";" -> "," across runs, plus safe bare-list wrapping only inside single w:t nodes.
  foreach ($p in $paras) {
    $stats.bracket_semicolons_replaced += (Normalize-BracketCitationsInParagraphRuns -Paragraph $p -Nsm $nsm)
    $sk = @()
    foreach ($t in @($p.SelectNodes(".//w:t", $nsm))) {
      $stats.bare_lists_wrapped += (Wrap-BareNumericListsInTextNode -TextNode $t -SkippedOut ([ref]$sk))
    }
    foreach ($x in $sk) { if ($stats.bare_lists_skipped.Count -lt 50) { $stats.bare_lists_skipped += $x } }
  }

  function Refresh-Paragraphs { $script:paras = $xml.SelectNodes("//w:body/w:p", $nsm) }

  $h441 = Find-BodyHeadingIndexByPrefix -Paras $paras -Nsm $nsm -Prefix "4.4.1" -StartAt 0
  if ($null -eq $h441) { throw "Could not find 4.4.1 in body." }
  $h45 = Find-BodyHeadingIndexByPrefix -Paras $paras -Nsm $nsm -Prefix "4.5" -StartAt $h441
  if ($null -eq $h45) { throw "Could not find 4.5 in body." }

  # Find all subsection headings in 4.4.*
  $subHeads = @()
  foreach ($pref in @("4.4.1", "4.4.2", "4.4.3", "4.4.4", "4.4.5")) {
    $ix = Find-BodyHeadingIndexByPrefix -Paras $paras -Nsm $nsm -Prefix $pref -StartAt $h441
    if ($null -ne $ix) { $subHeads += [pscustomobject]@{ pref = $pref; idx = $ix } }
  }
  $subHeads = $subHeads | Sort-Object idx

  # Capture caption/body styles for reporting.
  for ($i = $h441; $i -lt $h45; $i++) {
    $t = Get-ParagraphText -Paragraph $paras[$i] -Nsm $nsm
    if (-not $t) { continue }
    if ($t -match "^Ð Ð¸Ñ\\.") {
      $s = Paragraph-GetStyle -Paragraph $paras[$i] -Nsm $nsm
      if ($s -and -not ($stats.style_caption_used -contains $s)) { $stats.style_caption_used += $s }
    } elseif ($t -match "^(4\\.4\\.)") {
      continue
    } else {
      $s = Paragraph-GetStyle -Paragraph $paras[$i] -Nsm $nsm
      if ($s -and -not ($stats.style_body_used -contains $s)) { $stats.style_body_used += $s }
    }
  }

  for ($si = 0; $si -lt $subHeads.Count; $si++) {
    Refresh-Paragraphs
    $startIdx = ($subHeads[$si].idx)
    $endIdx = if ($si + 1 -lt $subHeads.Count) { ($subHeads[$si + 1].idx) } else { $h45 }

    # Determine body style for inserted legend (first normal paragraph style after heading).
    $bodyStyle = ""
    for ($j = $startIdx + 1; $j -lt $endIdx; $j++) {
      $pt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if (-not $pt) { continue }
      if ($pt -match "^(Ð Ð¸Ñ\\.|Ð Ð¸ÑÑƒÐ½Ð¾Ðº)") { continue }
      if ($pt -match "^4\\.4\\.") { continue }
      $bodyStyle = Paragraph-GetStyle -Paragraph $paras[$j] -Nsm $nsm
      if ($bodyStyle) { break }
    }
    if (-not $bodyStyle) { $bodyStyle = "a2" }

    # Remove any existing legend-like paragraph just under the heading.
    $windowEnd = [Math]::Min($paras.Count, $startIdx + 10)
    for ($j = $windowEnd - 1; $j -ge ($startIdx + 1); $j--) {
      $pt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if (-not $pt) { continue }
      if ($pt.StartsWith("Ð†Ð½Ñ‚ÐµÑ€Ð¿Ñ€ÐµÑ‚Ð°Ñ†Ñ–Ñ Ð¿Ð°Ð½ÐµÐ»ÐµÐ¹") -or $pt.StartsWith("Ð£ Ð²ÑÑ–Ñ… Ð½Ð°Ð²ÐµÐ´ÐµÐ½Ð¸Ñ… Ð´Ð°Ð»Ñ– Ñ–Ð»ÑŽÑÑ‚Ñ€Ð°Ñ†Ñ–ÑÑ… Ñ€Ð¾Ð·Ð´Ñ–Ð»Ñƒ 4.4")) {
        Remove-ParagraphAt -Paras $paras -Index $j
        $stats.legends_removed++
      }
    }
    Refresh-Paragraphs
    # Recompute indices after deletion.
    $startIdx = Find-BodyHeadingIndexByPrefix -Paras $paras -Nsm $nsm -Prefix $subHeads[$si].pref -StartAt $h441
    $endIdx = if ($si + 1 -lt $subHeads.Count) { Find-BodyHeadingIndexByPrefix -Paras $paras -Nsm $nsm -Prefix $subHeads[$si + 1].pref -StartAt $startIdx } else { Find-BodyHeadingIndexByPrefix -Paras $paras -Nsm $nsm -Prefix "4.5" -StartAt $startIdx }
    if ($null -eq $endIdx) { $endIdx = $paras.Count }

    $legendText = "Ð†Ð½Ñ‚ÐµÑ€Ð¿Ñ€ÐµÑ‚Ð°Ñ†Ñ–Ñ Ð¿Ð°Ð½ÐµÐ»ÐµÐ¹: Ð»Ñ–Ð²Ð° â€” Ð¾Ð¿Ñ‚Ð¸Ñ‡Ð½Ð¸Ð¹ Ð¿Ð¾Ñ‚Ñ–Ðº; Ñ†ÐµÐ½Ñ‚Ñ€Ð°Ð»ÑŒÐ½Ð° â€” ÐºÐ°Ñ€Ñ‚Ð° Ð³Ð»Ð¸Ð±Ð¸Ð½Ð¸; Ð¿Ñ€Ð°Ð²Ð° â€” Ð¾Ð±â€™Ñ”Ð´Ð½Ð°Ð½Ðµ Ð·Ð¾Ð±Ñ€Ð°Ð¶ÐµÐ½Ð½Ñ. Ð¡Ð»ÑƒÐ¶Ð±Ð¾Ð²Ñ– Ñ„Ñ€Ð°Ð·Ð¸ Ð¿Ñ€Ð¾ Ð´ÐµÑ‚ÐµÐºÑ†Ñ–ÑŽ (ÐºÑ–Ð»ÑŒÐºÑ–ÑÑ‚ÑŒ Ð²Ð¸ÑÐ²Ð»ÐµÐ½Ð¸Ñ… Ð¾Ð±â€™Ñ”ÐºÑ‚Ñ–Ð²) Ð½Ðµ Ð´ÑƒÐ±Ð»ÑŽÑŽÑ‚ÑŒÑÑ Ð² ÐºÐ¾Ð¶Ð½Ð¾Ð¼Ñƒ Ð¾Ð¿Ð¸ÑÑ–."
    Insert-ParagraphAfterIndex -Xml $xml -Nsm $nsm -Paras $paras -AfterIndex $startIdx -StyleVal $bodyStyle -Text $legendText
    $stats.legends_inserted++

    Refresh-Paragraphs
    # Refresh subsection bounds again after insertion.
    $startIdx = Find-BodyHeadingIndexByPrefix -Paras $paras -Nsm $nsm -Prefix $subHeads[$si].pref -StartAt $h441
    $endIdx = if ($si + 1 -lt $subHeads.Count) { Find-BodyHeadingIndexByPrefix -Paras $paras -Nsm $nsm -Prefix $subHeads[$si + 1].pref -StartAt $startIdx } else { Find-BodyHeadingIndexByPrefix -Paras $paras -Nsm $nsm -Prefix "4.5" -StartAt $startIdx }
    if ($null -eq $endIdx) { $endIdx = $paras.Count }

    # Gather caption indices within subsection.
    $capIdx = New-Object 'System.Collections.Generic.List[int]'
    $capInfo = @{}
    for ($j = $startIdx + 1; $j -lt $endIdx; $j++) {
      $pt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if ($pt -match "^Ð Ð¸Ñ\\.\\s*4\\.(\\d+)") {
        [void]$capIdx.Add($j)
        $capInfo[$j] = Get-FigureInfoFromCaption -CaptionText $pt
      }
    }
    if ($capIdx.Count -eq 0) { continue }

    # Build method-blocks by consecutive caption prefixes.
    $blocks = New-Object 'System.Collections.Generic.List[object]'
    $bStart = $capIdx[0]
    $bEnd = $capIdx[0]
    $bPrefix = $capInfo[$capIdx[0]].prefix
    $bFigs = @($capInfo[$capIdx[0]].fig)
    for ($ci = 1; $ci -lt $capIdx.Count; $ci++) {
      $cur = $capIdx[$ci]
      $curPrefix = $capInfo[$cur].prefix
      if ($curPrefix -eq $bPrefix -and ($cur - $bEnd) -le 10) {
        $bEnd = $cur
        $bFigs += $capInfo[$cur].fig
      } else {
        $blocks.Add([pscustomobject]@{ startCap = $bStart; endCap = $bEnd; prefix = $bPrefix; figs = $bFigs }) | Out-Null
        $bStart = $cur
        $bEnd = $cur
        $bPrefix = $curPrefix
        $bFigs = @($capInfo[$cur].fig)
      }
    }
    $blocks.Add([pscustomobject]@{ startCap = $bStart; endCap = $bEnd; prefix = $bPrefix; figs = $bFigs }) | Out-Null

    # Process blocks from bottom to top (stable deletion).
    for ($bi = $blocks.Count - 1; $bi -ge 0; $bi--) {
      Refresh-Paragraphs
      $block = $blocks[$bi]
      $capFirst = $block.startCap
      $capLast = $block.endCap

      # Span: after previous block's last caption (or after legend) up to last caption - 1.
      $spanStart = $startIdx + 2 # heading + legend
      if ($bi -gt 0) {
        $prevLast = $blocks[$bi - 1].endCap
        $spanStart = $prevLast + 1
      }
      $spanEnd = $capLast - 1
      if ($spanEnd -lt $spanStart) { continue }

      # Choose target paragraph to rewrite: first text-only paragraph in span.
      $targetIdx = $null
      $targetStyle = $bodyStyle
      for ($j = $spanStart; $j -le $spanEnd; $j++) {
        $pt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
        if (-not $pt) { continue }
        if (Paragraph-HasDrawing -Paragraph $paras[$j] -Nsm $nsm) { continue }
        if ($pt -match "^4\\.4\\.") { continue }
        if ($pt -match "^Ð Ð¸Ñ\\.") { continue }
        $targetIdx = $j
        $targetStyle = Paragraph-GetStyle -Paragraph $paras[$j] -Nsm $nsm
        if (-not $targetStyle) { $targetStyle = $bodyStyle }
        break
      }
      if ($null -eq $targetIdx) {
        # No existing paragraph to rewrite; insert a new one right before first caption.
        $insertAfter = $capFirst - 1
        Insert-ParagraphAfterIndex -Xml $xml -Nsm $nsm -Paras $paras -AfterIndex $insertAfter -StyleVal $bodyStyle -Text " "
        Refresh-Paragraphs
        $targetIdx = $insertAfter + 1
        $targetStyle = $bodyStyle
      }

      # Build source pool (text-only paragraphs in span), deleting standalone boilerplate paragraphs.
      $pool = New-Object 'System.Collections.Generic.List[string]'
      $toDelete = New-Object 'System.Collections.Generic.List[int]'
      for ($j = $spanStart; $j -le $spanEnd; $j++) {
        $pt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
        if (-not $pt) { continue }
        if ($pt -match "^Ð Ð¸Ñ\\.") { continue }
        if ($pt -match "^4\\.4\\.") { continue }
        if (Paragraph-HasDrawing -Paragraph $paras[$j] -Nsm $nsm) { continue }

        $low = $pt.ToLowerInvariant()
        if ($low.StartsWith("Ð»Ñ–Ð²Ð° Ð¿Ð°Ð½ÐµÐ»ÑŒ") -or $low.StartsWith("Ñ†ÐµÐ½Ñ‚Ñ€Ð°Ð»ÑŒÐ½Ð° Ð¿Ð°Ð½ÐµÐ»ÑŒ") -or $low.StartsWith("Ð¿Ñ€Ð°Ð²Ð° Ð¿Ð°Ð½ÐµÐ»ÑŒ")) {
          $toDelete.Add($j) | Out-Null
          $stats.panel_paras_deleted++
          continue
        }
        if ($low.StartsWith("Ð´ÐµÑ‚ÐµÐºÑ†Ñ–Ñ") -or $low.StartsWith("Ð²Ð¸ÑÐ²Ð»ÐµÐ½Ð¾") -or $low.StartsWith("Ñƒ Ð½Ð¸Ð¶Ð½Ñ–Ð¹ Ñ‡Ð°ÑÑ‚Ð¸Ð½Ñ–")) {
          $toDelete.Add($j) | Out-Null
          $stats.detection_paras_deleted++
          continue
        }
        $pool.Add($pt) | Out-Null
        if ($j -ne $targetIdx) { $toDelete.Add($j) | Out-Null }
      }

      # Create 4-sentence paragraph from pool (use your wording, remove repeats).
      $cites = @()
      $sentOut = New-Object 'System.Collections.Generic.List[string]'
      $seen = New-Object 'System.Collections.Generic.HashSet[string]'
      foreach ($ptext in $pool) {
        $clean = Clean-BoilerplateFromText -Text $ptext
        foreach ($s in (Split-Sentences -Text $clean)) {
          $s2 = Extract-AndStripCitations -Sentence $s -Cites ([ref]$cites)
          $key = ($s2.ToLowerInvariant() -replace "\s+", " ").Trim()
          if (-not $key) { continue }
          if ($seen.Contains($key)) { continue }
          [void]$seen.Add($key)
          [void]$sentOut.Add($s2)
          if ($sentOut.Count -ge 3) { break }
        }
        if ($sentOut.Count -ge 3) { break }
      }
      while ($sentOut.Count -lt 3) { [void]$sentOut.Add(" ") }

      $figNums = @($block.figs | Sort-Object)
      $n = $figNums[0]
      $m = $figNums[$figNums.Count - 1]
      $range = if ($n -eq $m) { "4.$n" } else { "4.$nâ€“4.$m" }
      $citeMerged = Merge-CitationBlocks -Blocks ($cites | Select-Object -Unique)
      $s4 = "ÐÐ° Ñ€Ð¸Ñ. $range Ð½Ð°Ð²ÐµÐ´ÐµÐ½Ð¾ Ñ€ÐµÐ·ÑƒÐ»ÑŒÑ‚Ð°Ñ‚Ð¸ Ð´Ð»Ñ $($block.prefix)."
      if ($citeMerged) { $s4 = $s4.TrimEnd('.') + " " + $citeMerged + "." }

      $final = (($sentOut[0].TrimEnd('.') + ".").Trim()) + " " +
               (($sentOut[1].TrimEnd('.') + ".").Trim()) + " " +
               (($sentOut[2].TrimEnd('.') + ".").Trim()) + " " +
               $s4
      $final = ($final -replace "\s{2,}", " ").Trim()

      # Rewrite target paragraph, delete other text paragraphs in span.
      Paragraph-SetStyle -Paragraph $paras[$targetIdx] -Nsm $nsm -StyleVal $targetStyle
      Rewrite-TextOnlyParagraph -Paragraph $paras[$targetIdx] -Nsm $nsm -Text $final
      $stats.rewritten_blocks++
      $stats.rewritten_paras++

      # Delete other paragraphs (from bottom to top).
      foreach ($del in ($toDelete | Sort-Object -Descending)) {
        if ($del -eq $targetIdx) { continue }
        Remove-ParagraphAt -Paras $paras -Index $del
        $stats.deleted_text_paras++
        Refresh-Paragraphs
      }
    }
  }

  Save-XmlDocumentUtf8NoBom -Xml $xml -Path $docXmlPath
  Compress-Docx -SourceDir $workDir -OutDocxPath $OutputDocx

  Ensure-ParentDir -Path $ReportPath
  ($stats | ConvertTo-Json -Depth 8) | Out-File -LiteralPath $ReportPath -Encoding UTF8
} finally {
  if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
  }
}

