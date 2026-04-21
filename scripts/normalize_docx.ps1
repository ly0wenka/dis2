param(
  [Parameter(Mandatory = $true)][string]$InputDocx,
  [Parameter(Mandatory = $true)][string]$OutputDocx,
  [Parameter(Mandatory = $false)][ValidateSet("Whole")][string]$Scope = "Whole",
  [Parameter(Mandatory = $false)][switch]$Fix44,
  [Parameter(Mandatory = $false)][switch]$Remove44Formulas,
  [Parameter(Mandatory = $false)][switch]$NormalizeBracketCitations,
  [Parameter(Mandatory = $false)][switch]$RemoveRefTokens,
  [Parameter(Mandatory = $false)][ValidateSet("KeepTextRemoveStrike", "DeleteText")][string]$StrikeHandling = "KeepTextRemoveStrike",
  [Parameter(Mandatory = $false)][string]$ReportPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

. "$PSScriptRoot\\docx_lib.ps1"

function Ensure-ParentDir {
  param([Parameter(Mandatory = $true)][string]$Path)
  $parent = Split-Path -Parent $Path
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent | Out-Null
  }
}

$workRoot = Join-Path (Split-Path -Parent $OutputDocx) "tmp\\work"
$workDir = Join-Path $workRoot ("docx_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

try {
  Expand-Docx -DocxPath $InputDocx -DestinationDir $workDir

  $docXmlPath = Join-Path $workDir "word\\document.xml"
  $xml = Load-XmlDocument -Path $docXmlPath
  $nsm = New-WordNamespaceManager -Xml $xml

  $stats = [ordered]@{
    input = (Resolve-Path -LiteralPath $InputDocx).Path
    output = (Resolve-Path -LiteralPath (Split-Path -Parent $OutputDocx)).Path + "\\" + (Split-Path -Leaf $OutputDocx)
    removed_ref_tokens = 0
    normalized_citations = 0
    strike_runs_processed = 0
    strike_runs_deleted = 0
    strike_marks_removed = 0
    fix44_para_start = $null
    fix44_para_end = $null
    fix44_math_paras_removed = 0
    fix44_list_paras_flattened = 0
    fix44_short_paras_merged = 0
    fix44_panel_dupes_removed = 0
    fix44_detection_paras_removed = 0
    fix44_detection_text_removed = 0
    fix44_ref_paras_removed = 0
    fix44_ref_inserted = 0
  }

  # Whole document text transforms
  foreach ($t in $xml.SelectNodes("//w:t", $nsm)) {
    $orig = $t.InnerText
    $new = $orig
    if ($RemoveRefTokens) { $new = Remove-RefTokensInText $new }
    if ($NormalizeBracketCitations) { $new = Normalize-BracketCitationsInText $new }
    if ($new -ne $orig) {
      if ($RemoveRefTokens -and ($orig -match "REF\s+_Ref") -and ($new -notmatch "REF\s+_Ref")) { $stats.removed_ref_tokens++ }
      if ($NormalizeBracketCitations -and ($orig -match "\[[^\]]*;[^\]]*\]") -and ($new -notmatch "\[[^\]]*;[^\]]*\]")) { $stats.normalized_citations++ }
      $t.InnerText = $new
    }
  }

  # Whole document strikethrough
  $runs = $xml.SelectNodes("//w:r", $nsm)
  for ($i = $runs.Count - 1; $i -ge 0; $i--) {
    $r = $runs[$i]
    $rPr = $r.SelectSingleNode("./w:rPr", $nsm)
    if (-not $rPr) { continue }
    $strike = $rPr.SelectSingleNode("./w:strike | ./w:dstrike", $nsm)
    if (-not $strike) { continue }

    $stats.strike_runs_processed++
    if ($StrikeHandling -eq "DeleteText") {
      [void]$r.ParentNode.RemoveChild($r)
      $stats.strike_runs_deleted++
      continue
    }
    foreach ($node in @($rPr.SelectNodes("./w:strike | ./w:dstrike", $nsm))) {
      [void]$rPr.RemoveChild($node)
      $stats.strike_marks_removed++
    }
  }

  if ($Fix44) {
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)

    function Find-ParaIndexByRegex {
      param([Parameter(Mandatory = $true)][string]$Pattern, [Parameter(Mandatory = $false)][int]$StartAt = 0)
      for ($j = $StartAt; $j -lt $paras.Count; $j++) {
        $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
        if ($txt -match $Pattern) { return $j }
      }
      return $null
    }

    function Find-ParaIndexByPrefix {
      param([Parameter(Mandatory = $true)][string]$Prefix, [Parameter(Mandatory = $false)][int]$StartAt = 0)
      for ($j = $StartAt; $j -lt $paras.Count; $j++) {
        $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
        if ($txt -and $txt.StartsWith($Prefix)) { return $j }
      }
      return $null
    }

    # Avoid relying on Cyrillic anchors: skip front matter/TOC by starting late.
    $startSearchAt = 400
    $idx441 = Find-ParaIndexByPrefix -Prefix "4.4.1" -StartAt $startSearchAt
    if ($null -eq $idx441) { throw "Could not locate heading 4.4.1 in document body." }
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    if ($null -eq $idx445) { throw "Could not locate heading 4.4.5 in document body." }
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    $stats.fix44_para_start = $idx441
    $stats.fix44_para_end = $idxAfter445

    function Insert-ParagraphAfter {
      param(
        [Parameter(Mandatory = $true)][int]$AfterIndex,
        [Parameter(Mandatory = $true)][string]$Text
      )
      $pNew = $xml.CreateElement("w", "p", $nsm.LookupNamespace("w"))
      $r = $xml.CreateElement("w", "r", $nsm.LookupNamespace("w"))
      $t = $xml.CreateElement("w", "t", $nsm.LookupNamespace("w"))
      $t.SetAttribute("xml:space", "preserve")
      $t.InnerText = $Text
      [void]$r.AppendChild($t)
      [void]$pNew.AppendChild($r)
      [void]$paras[$AfterIndex].ParentNode.InsertAfter($pNew, $paras[$AfterIndex])
      $stats.fix44_ref_inserted++
    }

    # Normalize repeated DETR detection boilerplate and reference sentences:
    # - Remove duplicated consolidated note right under 4.4.1 (the duplicates accumulate across runs),
    # - Remove standalone detection lines across 4.4.1–4.4.5,
    # - Re-insert exactly one consolidated note under 4.4.1.
    $noteSig = "блок\s+детекції\s+моделі\s+DETR"
    $noteText = "У всіх наведених далі ілюстраціях розділу 4.4 блок детекції моделі DETR відображає кількість виявлених у кадрі динамічних об’єктів (наприклад, «6 об’єктів виявлено»). Це значення використовується для подальшого зв’язування оцінених параметрів руху та глибини з конкретними об’єктами; тому однакові службові пояснення не дублюються для кожного рисунка окремо."

    # Remove any existing consolidated note paragraphs in a small window below 4.4.1 (robust even if range bounds drift).
    $windowEnd = [Math]::Min($paras.Count, $idx441 + 25)
    for ($j = $windowEnd - 1; $j -ge ($idx441 + 1); $j--) {
      $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if ($txt -match $noteSig) {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_detection_paras_removed++
      }
    }
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)

    # Remove standalone detection lines across 4.4.1–4.4.5
    for ($j = $idxAfter445 - 1; $j -ge $idx441; $j--) {
      $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if (-not $txt) { continue }
      $lower = $txt.ToLowerInvariant()
      if ($txt -match "об.?єктів\s+виявлено" -or $lower.StartsWith("детекція об")) {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_detection_paras_removed++
      }
    }
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)

    # Insert one consolidated note once (directly below 4.4.1)
    Insert-ParagraphAfter -AfterIndex $idx441 -Text $noteText
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)

    # Remove any existing reference lines (4.4.4 and 4.4.5) so we can re-insert cleanly.
    for ($j = $idxAfter445 - 1; $j -ge $idx441; $j--) {
      $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if (-not $txt) { continue }
      if ($txt -match "^Параметри\s+руху\s+обчиснюються\s+за\s+постановкою") {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_ref_paras_removed++
        continue
      }
      if ($txt -match "^Параметри\s+руху\s+для\s+локальної\s+моделі") {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_ref_paras_removed++
      }
    }
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    # Refresh indices after removals/insertions
    $idx441 = Find-ParaIndexByPrefix -Prefix "4.4.1" -StartAt $startSearchAt
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    # Insert reference sentences under 4.4.4 and 4.4.5 headings (exactly once).
    $idx444 = Find-ParaIndexByPrefix -Prefix "4.4.4" -StartAt $idx441
    if ($null -ne $idx444) {
      Insert-ParagraphAfter -AfterIndex $idx444 -Text "Параметри руху обчиснюються за постановкою та ітераційною схемою розд. 2, формули (2.1)-(2.10)."
    }
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    if ($null -ne $idx445) {
      Insert-ParagraphAfter -AfterIndex $idx445 -Text "Параметри руху для локальної моделі оцінювання руху наведено в розд. 2, формули (2.11)-(2.24)."
    }
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt $idx441
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    if ($Remove44Formulas) {
      $toRemove = New-Object 'System.Collections.Generic.HashSet[int]'
      for ($j = $idx441; $j -lt $idxAfter445; $j++) {
        if (Paragraph-HasMath -Paragraph $paras[$j] -Nsm $nsm) {
          [void]$toRemove.Add($j)
          if ($j -gt $idx441) {
            $prevTxt = Get-ParagraphText -Paragraph $paras[$j - 1] -Nsm $nsm
            if (Paragraph-IsOnlyEquationNumberText $prevTxt) { [void]$toRemove.Add($j - 1) }
          }
          if ($j + 1 -lt $idxAfter445) {
            $nextTxt = Get-ParagraphText -Paragraph $paras[$j + 1] -Nsm $nsm
            if (Paragraph-IsOnlyEquationNumberText $nextTxt) { [void]$toRemove.Add($j + 1) }
          }
        }
      }
      for ($j = $idxAfter445 - 1; $j -ge $idx441; $j--) {
        if ($toRemove.Contains($j)) {
          [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
          $stats.fix44_math_paras_removed++
        }
      }
    }

    # Refresh
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    $idx441 = Find-ParaIndexByPrefix -Prefix "4.4.1" -StartAt $startSearchAt
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt $idx441
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    function Para-HasNumPr {
      param([Parameter(Mandatory = $true)][System.Xml.XmlNode]$P)
      return ($P.SelectSingleNode("./w:pPr/w:numPr", $nsm) -ne $null)
    }
    function Remove-NumPr {
      param([Parameter(Mandatory = $true)][System.Xml.XmlNode]$P)
      $numPr = $P.SelectSingleNode("./w:pPr/w:numPr", $nsm)
      if ($numPr) { [void]$numPr.ParentNode.RemoveChild($numPr) }
    }
    function Append-Runs {
      param(
        [Parameter(Mandatory = $true)][System.Xml.XmlNode]$TargetP,
        [Parameter(Mandatory = $true)][System.Xml.XmlNode]$SourceP,
        [Parameter(Mandatory = $true)][string]$Separator
      )
      $rSep = $xml.CreateElement("w", "r", $nsm.LookupNamespace("w"))
      $tSep = $xml.CreateElement("w", "t", $nsm.LookupNamespace("w"))
      $tSep.SetAttribute("xml:space", "preserve")
      $tSep.InnerText = $Separator
      [void]$rSep.AppendChild($tSep)
      [void]$TargetP.AppendChild($rSep)

      foreach ($r in @($SourceP.SelectNodes("./w:r", $nsm))) {
        [void]$TargetP.AppendChild($xml.ImportNode($r, $true))
      }
    }

    $panelSeen = @{
      left = $false
      center = $false
      right = $false
    }
    function Reset-PanelSeen { $panelSeen.left = $false; $panelSeen.center = $false; $panelSeen.right = $false }
    Reset-PanelSeen

    $j = $idx441
    while ($j -lt $idxAfter445) {
      $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if (-not $txt) { $j++; continue }

      if (Paragraph-IsHeadingText $txt -or Paragraph-IsCaptionText $txt) {
        Reset-PanelSeen
        $j++; continue
      }

      $lower = $txt.ToLowerInvariant()
      $key = $null
      if ($lower -like "ліва панель*") { $key = "left" }
      elseif ($lower -like "центральна панель*") { $key = "center" }
      elseif ($lower -like "права панель*") { $key = "right" }

      if ($key) {
        if ($panelSeen[$key]) {
          [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
          $stats.fix44_panel_dupes_removed++
          $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
          $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt $idx441
          if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }
          continue
        }
        $panelSeen[$key] = $true
      }

      $isListish = ($txt -like "- *") -or (Para-HasNumPr -P $paras[$j])
      if ($isListish) {
        Remove-NumPr -P $paras[$j]
        $firstT = $paras[$j].SelectSingleNode(".//w:t", $nsm)
        if ($firstT -and $firstT.InnerText -match "^\s*-\s+") {
          $firstT.InnerText = ($firstT.InnerText -replace "^\s*-\s+", "")
        }
        $stats.fix44_list_paras_flattened++
      }

      $canMerge = ($txt.Length -lt 140) -and ($j -gt $idx441)
      if ($canMerge) {
        $prevTxt = Get-ParagraphText -Paragraph $paras[$j - 1] -Nsm $nsm
        $prevBarrier = (-not $prevTxt) -or (Paragraph-IsHeadingText $prevTxt) -or (Paragraph-IsCaptionText $prevTxt)
        if (-not $prevBarrier) {
          Append-Runs -TargetP $paras[$j - 1] -SourceP $paras[$j] -Separator " "
          [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
          $stats.fix44_short_paras_merged++
          $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
          $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt $idx441
          if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }
          continue
        }
      }
      $j++
    }

    # Remove remaining embedded DETR detection boilerplate inside paragraphs (when it was merged into longer text).
    for ($j = $idx441; $j -lt $idxAfter445; $j++) {
      $p = $paras[$j]
      $tNodes = @($p.SelectNodes(".//w:t", $nsm))
      if ($tNodes.Count -eq 0) { continue }

      $full = Get-ParagraphText -Paragraph $p -Nsm $nsm
      if (-not $full) { continue }
      if ($full -match "блок\s+детекції\s+моделі\s+DETR") { continue } # keep consolidated note

      $newFull = $full
      $newFull = $newFull -replace "Детекц[іяі][^\\.]{0,260}DETR[^\\.]{0,260}\\.?\s*", ""
      $newFull = $newFull -replace "Детекц[іяі][^\\.]{0,80}нижній частині[^\\.]{0,360}виявлено[^\\.]{0,200}\\.?\s*", ""
      $newFull = $newFull -replace "У нижній частині[^\\.]{0,360}виявлено[^\\.]{0,200}\\.?\s*", ""
      $newFull = $newFull -replace "«\s*\d+\s+об.?єктів\s+виявлено\s*»\s*", ""
      $newFull = $newFull -replace "\s{2,}", " "
      $newFull = $newFull.Trim()

      if ($newFull -ne $full) {
        $first = $p.SelectSingleNode(".//w:t", $nsm)
        if ($first) {
          $first.InnerText = $newFull
          for ($k = $tNodes.Count - 1; $k -ge 1; $k--) {
            [void]$tNodes[$k].ParentNode.RemoveChild($tNodes[$k])
          }
          $stats.fix44_detection_text_removed++
        }
      }
    }
  }

  Save-XmlDocumentUtf8NoBom -Xml $xml -Path $docXmlPath
  Compress-Docx -SourceDir $workDir -OutDocxPath $OutputDocx

  if ($ReportPath) {
    Ensure-ParentDir -Path $ReportPath
    ($stats | ConvertTo-Json -Depth 6) | Out-File -LiteralPath $ReportPath -Encoding UTF8
  }
} finally {
  if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
  }
}
