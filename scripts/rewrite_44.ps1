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

function Paragraph-HasBreak {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
  )
  if ($Paragraph.SelectSingleNode(".//w:br[@w:type='page'] | .//w:lastRenderedPageBreak | .//w:pPr/w:sectPr", $Nsm)) {
    return $true
  }
  return $false
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

function Paragraph-RemoveNumbering {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
  )
  $numPr = $Paragraph.SelectSingleNode("./w:pPr/w:numPr", $Nsm)
  if ($numPr) { [void]$numPr.ParentNode.RemoveChild($numPr) }
}

function Rewrite-TextOnlyParagraph {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm,
    [Parameter(Mandatory = $true)][string]$Text
  )
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

function New-ParagraphWithStyleAndText {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Xml,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm,
    [Parameter(Mandatory = $false)][AllowEmptyString()][string]$StyleVal,
    [Parameter(Mandatory = $true)][string]$Text
  )
  $pNew = $Xml.CreateElement("w", "p", $Nsm.LookupNamespace("w"))
  $pPr = $Xml.CreateElement("w", "pPr", $Nsm.LookupNamespace("w"))
  if ($StyleVal) {
    $pStyle = $Xml.CreateElement("w", "pStyle", $Nsm.LookupNamespace("w"))
    $pStyle.SetAttribute("val", $Nsm.LookupNamespace("w"), $StyleVal)
    [void]$pPr.AppendChild($pStyle)
  }
  [void]$pNew.AppendChild($pPr)
  $null = Rewrite-TextOnlyParagraph -Paragraph $pNew -Nsm $Nsm -Text $Text
  return $pNew
}

function Find-BodyHeadingIndex {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode[]]$Paras,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm,
    [Parameter(Mandatory = $true)][string]$Needle
  )
  for ($i = 0; $i -lt $Paras.Count; $i++) {
    $t = Get-ParagraphText -Paragraph $Paras[$i] -Nsm $Nsm
    if ($t.StartsWith($Needle)) {
      $max = [Math]::Min($Paras.Count - 1, $i + 200)
      for ($j = $i; $j -le $max; $j++) {
        if (Paragraph-HasDrawing -Paragraph $Paras[$j] -Nsm $Nsm) { return $i }
      }
    }
  }
  return -1
}

function Find-NextHeadingIndexByRegex {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode[]]$Paras,
    [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm,
    [Parameter(Mandatory = $true)][int]$StartIndex,
    [Parameter(Mandatory = $true)][string]$Pattern
  )
  for ($i = $StartIndex; $i -lt $Paras.Count; $i++) {
    $t = Get-ParagraphText -Paragraph $Paras[$i] -Nsm $Nsm
    if ($t -match $Pattern) { return $i }
  }
  return -1
}

function Try-ParseFigureCaption {
  param([Parameter(Mandatory = $true)][string]$Text)
  # Example: "Рис. 4.62 — Farneback_..."
  $m = [regex]::Match($Text, "^Рис\.\s*4\.(\d+)\s*[\u2014\u2013-]\s*(.+)$")
  if (-not $m.Success) { return $null }
  $figNo = [int]$m.Groups[1].Value
  $tail = $m.Groups[2].Value.Trim()
  $prefix = $tail
  $us = $tail.IndexOf("_")
  if ($us -gt 0) { $prefix = $tail.Substring(0, $us) }
  else {
    $sp = $tail.IndexOf(" ")
    if ($sp -gt 0) { $prefix = $tail.Substring(0, $sp) }
  }
  $prefix = $prefix.Trim()
  if (-not $prefix) { $prefix = "Method" }
  return [pscustomobject]@{ fig = $figNo; prefix = $prefix; tail = $tail }
}

function Is-BoilerplateParagraphText {
  param([Parameter(Mandatory = $true)][string]$Text)
  $s = $Text.Trim().ToLowerInvariant()
  if (-not $s) { return $false }
  if ($s.StartsWith("ліва панель") -or $s.StartsWith("центральна панель") -or $s.StartsWith("права панель")) { return $true }
  if ($s.StartsWith("на лівій панелі") -or $s.StartsWith("на центральній панелі") -or $s.StartsWith("на правій панелі")) { return $true }
  if ($s.StartsWith("детекція") -or $s.StartsWith("виявлено") -or $s.StartsWith("у нижній частині")) { return $true }
  if ($s -match "\bobjects\s+detected\b") { return $true }
  return $false
}

function Clean-BoilerplateFromText {
  param([Parameter(Mandatory = $true)][string]$Text)
  $rx = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
  $t = $Text
  $t = [regex]::Replace($t, "Детекція\s+об’єктів\.?\s*", "", $rx)
  $t = [regex]::Replace($t, "Виявлено\s+\d+\s+об.?’?єктів\s+моделлю\s+DETR\.?\s*", "", $rx)
  $t = [regex]::Replace($t, "Детекція\s+\d+\s+об.?’?єктів\s+моделлю\s+DETR\.?\s*", "", $rx)
  $t = [regex]::Replace($t, "У\s+нижній\s+частині.*?\.\s*", "", $rx)
  $t = [regex]::Replace($t, "Ліва\s+панель.*?\.\s*", "", $rx)
  $t = [regex]::Replace($t, "Центральна\s+панель.*?\.\s*", "", $rx)
  $t = [regex]::Replace($t, "Права\s+панель.*?\.\s*", "", $rx)
  $t = [regex]::Replace($t, "На\s+лівій\s+панелі.*?\.\s*", "", $rx)
  $t = [regex]::Replace($t, "На\s+центральній\s+панелі.*?\.\s*", "", $rx)
  $t = [regex]::Replace($t, "На\s+правій\s+панелі.*?\.\s*", "", $rx)
  $t = [regex]::Replace($t, "\b\d+\s+objects\s+detected\b\.?\s*", "", $rx)
  $t = ($t -replace "\s{2,}", " ").Trim()
  return $t
}

function Fix-BareCitationListsInString {
  param(
    [Parameter(Mandatory = $true)][string]$Text,
    [Parameter(Mandatory = $true)][ref]$Stats
  )
  $s = $Text
  $pattern = "(?<!\[)(?<list>\b\d+(?:\s*;\s*\d+){2,})(?=\s*[\.,;:\)\]]|$)"
  while ($true) {
    $m = [regex]::Match($s, $pattern)
    if (-not $m.Success) { break }
    $list = $m.Groups["list"].Value
    if ($list -match "\.") {
      $Stats.Value.citation_lists_skipped_mixed++
      break
    }
    $nums = ($list -split ";") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    if (($nums | Where-Object { $_ -notmatch "^\d+$" }).Count -gt 0) {
      $Stats.Value.citation_lists_skipped_mixed++
      break
    }
    $repl = "[" + (($nums | ForEach-Object { [int]$_ }) -join ", ") + "]"
    $s = $s.Substring(0, $m.Index) + $repl + $s.Substring($m.Index + $m.Length)
    $Stats.Value.citation_lists_wrapped++
  }
  return $s
}

function Extract-UniqueCitations {
  param([Parameter(Mandatory = $true)][string]$Text)
  $cites = New-Object System.Collections.Generic.List[string]
  foreach ($m in [regex]::Matches($Text, "\[(?<inner>[^\]]+)\]")) {
    $inner = $m.Groups["inner"].Value.Trim()
    if (-not $inner) { continue }
    $inner = ($inner -replace "\s*;\s*", ", ") -replace "\s{2,}", " "
    # keep as-is (may include pages "с. 21")
    if (-not ($cites.Contains($inner))) { $cites.Add($inner) | Out-Null }
  }
  return $cites
}

function Pick-Sentences {
  param([Parameter(Mandatory = $true)][string]$Text)
  $parts = [regex]::Split($Text, "(?<=[\.\!\?])\s+")
  $out = New-Object System.Collections.Generic.List[string]
  foreach ($p in $parts) {
    $s = $p.Trim()
    if (-not $s) { continue }
    if (Is-BoilerplateParagraphText -Text $s) { continue }
    $s = Clean-BoilerplateFromText -Text $s
    if (-not $s) { continue }
    if (-not ($s.EndsWith(".")) -and -not ($s.EndsWith("!")) -and -not ($s.EndsWith("?"))) { $s += "." }
    if (-not ($out.Contains($s))) { $out.Add($s) | Out-Null }
    if ($out.Count -ge 8) { break }
  }
  return $out
}

$stats = [ordered]@{
  input = $InputDocx
  output = $OutputDocx
  inserted_legends = 0
  inserted_summaries = 0
  deleted_boilerplate_paras = 0
  citation_semicolons_normalized = 0
  citation_lists_wrapped = 0
  citation_lists_skipped_mixed = 0
  sections_processed = @()
}

$workDir = Join-Path $env:TEMP ("docx_work_" + [guid]::NewGuid().ToString("N"))
try {
  Expand-Docx -DocxPath $InputDocx -DestinationDir $workDir
  $docXmlPath = Join-Path $workDir "word\\document.xml"
  if (-not (Test-Path -LiteralPath $docXmlPath)) { throw "Missing word/document.xml in DOCX." }

  $xml = Load-XmlDocument -Path $docXmlPath
  $nsm = New-WordNamespaceManager -Xml $xml
  $body = $xml.SelectSingleNode("/w:document/w:body", $nsm)
  if (-not $body) { throw "Missing w:body." }

  $paras = @($body.SelectNodes("./w:p", $nsm))

  $h441 = Find-BodyHeadingIndex -Paras $paras -Nsm $nsm -Needle "4.4.1"
  $h442 = Find-BodyHeadingIndex -Paras $paras -Nsm $nsm -Needle "4.4.2"
  $h443 = Find-BodyHeadingIndex -Paras $paras -Nsm $nsm -Needle "4.4.3"
  $h444 = Find-BodyHeadingIndex -Paras $paras -Nsm $nsm -Needle "4.4.4"
  $h445 = Find-BodyHeadingIndex -Paras $paras -Nsm $nsm -Needle "4.4.5"
  if ($h441 -lt 0 -or $h442 -lt 0 -or $h443 -lt 0 -or $h444 -lt 0 -or $h445 -lt 0) {
    throw "Could not locate body headings 4.4.1-4.4.5 (skipping TOC)."
  }

  $h45 = Find-NextHeadingIndexByRegex -Paras $paras -Nsm $nsm -StartIndex ($h445 + 1) -Pattern "^4\.5(\s|$)"
  if ($h45 -lt 0) { throw "Could not find next heading 4.5 after 4.4.5." }

  $sections = @(
    @{ name = "4.4.1"; start = $h441; end = $h442 - 1 },
    @{ name = "4.4.2"; start = $h442; end = $h443 - 1 },
    @{ name = "4.4.3"; start = $h443; end = $h444 - 1 },
    @{ name = "4.4.4"; start = $h444; end = $h445 - 1 },
    @{ name = "4.4.5"; start = $h445; end = $h45 - 1 }
  )

  foreach ($sec in $sections) {
    # Refresh paragraph list each section, because we insert/delete.
    $paras = @($body.SelectNodes("./w:p", $nsm))

    $startIdx = [int]$sec.start
    $endIdx = [int]$sec.end
    $secName = [string]$sec.name

    if ($endIdx -le $startIdx -or $startIdx -lt 0 -or $endIdx -ge $paras.Count) { continue }

    # Pick a body style for inserted paragraphs from the first non-empty paragraph after heading.
    $bodyStyle = ""
    for ($k = $startIdx + 1; $k -le [Math]::Min($endIdx, $startIdx + 30); $k++) {
      $txtK = Get-ParagraphText -Paragraph $paras[$k] -Nsm $nsm
      if (-not $txtK) { continue }
      if (Paragraph-HasDrawing -Paragraph $paras[$k] -Nsm $nsm) { continue }
      if (Paragraph-HasBreak -Paragraph $paras[$k] -Nsm $nsm) { continue }
      $bodyStyle = Paragraph-GetStyle -Paragraph $paras[$k] -Nsm $nsm
      break
    }

    # Insert legend (once per subsection) if not already present right after heading.
    $nextTxt = Get-ParagraphText -Paragraph $paras[$startIdx + 1] -Nsm $nsm
    if (-not $nextTxt.StartsWith("Інтерпретація панелей:")) {
      $legendText = "Інтерпретація панелей: ліва — оптичний потік; центральна — карта глибини; права — об’єднане зображення. Службові фрази про детекцію (кількість виявлених об’єктів) не повторюються для кожного рисунка."
      $pLegend = New-ParagraphWithStyleAndText -Xml $xml -Nsm $nsm -StyleVal $bodyStyle -Text $legendText
      [void]$body.InsertAfter($pLegend, $paras[$startIdx])
      $stats.inserted_legends++
      $paras = @($body.SelectNodes("./w:p", $nsm))
      $endIdx++
    }

    # Recompute because we may have inserted.
    $paras = @($body.SelectNodes("./w:p", $nsm))

    # Find figure captions in this subsection.
    $captions = New-Object System.Collections.Generic.List[object]
    for ($i = $startIdx + 1; $i -le $endIdx; $i++) {
      $t = Get-ParagraphText -Paragraph $paras[$i] -Nsm $nsm
      if (-not $t) { continue }
      $info = Try-ParseFigureCaption -Text $t
      if ($info -ne $null) {
        $captions.Add([pscustomobject]@{ idx = $i; fig = $info.fig; prefix = $info.prefix; text = $t }) | Out-Null
      }
    }

    if ($captions.Count -gt 0) {
      # Group consecutive captions by prefix.
      $blocks = New-Object System.Collections.Generic.List[object]
      $cur = $null
      foreach ($c in $captions) {
        if ($cur -eq $null -or $cur.prefix -ne $c.prefix) {
          if ($cur -ne $null) { $blocks.Add($cur) | Out-Null }
          $cur = [pscustomobject]@{ prefix = $c.prefix; captions = New-Object System.Collections.Generic.List[object] }
        }
        $cur.captions.Add($c) | Out-Null
      }
      if ($cur -ne $null) { $blocks.Add($cur) | Out-Null }

      foreach ($b in $blocks) {
        $firstCap = $b.captions[0]
        $lastCap = $b.captions[$b.captions.Count - 1]

        # Insert summary paragraph BEFORE the first drawing that belongs to the first caption (so it's above the figure).
        $insertBeforeIdx = $firstCap.idx
        for ($j = $firstCap.idx; $j -ge ($startIdx + 1); $j--) {
          if (Paragraph-HasDrawing -Paragraph $paras[$j] -Nsm $nsm) {
            $insertBeforeIdx = $j
            break
          }
          # Stop if we hit another caption/heading.
          $tJ = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
          if ($tJ -match "^4\.4\.[1-5]\b" -or $tJ -match "^Рис\.") { break }
        }

        # Avoid inserting duplicates if a summary already exists nearby.
        $already = $false
        for ($chk = [Math]::Max($startIdx + 1, $insertBeforeIdx - 5); $chk -lt $insertBeforeIdx; $chk++) {
          $tChk = Get-ParagraphText -Paragraph $paras[$chk] -Nsm $nsm
          if ($tChk -match "^На\s+рис\.") { $already = $true; break }
        }
        if ($already) { continue }

        # Build a pool of nearby text (do not modify existing paragraphs; only use their sentences).
        $pool = New-Object System.Text.StringBuilder
        $poolCites = New-Object System.Collections.Generic.List[string]
        $winStart = [Math]::Max($startIdx + 1, $insertBeforeIdx - 35)
        for ($p = $winStart; $p -lt $insertBeforeIdx; $p++) {
          $pt = Get-ParagraphText -Paragraph $paras[$p] -Nsm $nsm
          if (-not $pt) { continue }
          if (Paragraph-HasDrawing -Paragraph $paras[$p] -Nsm $nsm) { continue }
          if (Paragraph-HasBreak -Paragraph $paras[$p] -Nsm $nsm) { continue }
          if ($pt -match "^4\.4\.[1-5]\b") { continue }
          if ($pt -match "^Рис\.") { continue }
          if (Is-BoilerplateParagraphText -Text $pt) { continue }

          $pt2 = Normalize-BracketCitationsInText -Text @($pt)
          $pt2 = Remove-RefTokensInText -Text @($pt2)
          $pt2 = Fix-BareCitationListsInString -Text $pt2 -Stats ([ref]$stats)

          $cites = Extract-UniqueCitations -Text $pt2
          foreach ($ci in $cites) {
            if (-not ($poolCites.Contains($ci))) { $poolCites.Add($ci) | Out-Null }
          }

          [void]$pool.Append(" ")
          [void]$pool.Append($pt2)
        }

        $poolText = ($pool.ToString() -replace "\s{2,}", " ").Trim()
        $sentences = Pick-Sentences -Text $poolText

        # Ensure we have 3 content sentences (safe fillers if needed).
        $s1 = if ($sentences.Count -ge 1) { $sentences[0] } else { "У підрозділі наведено узагальнений опис експериментальних результатів." }
        $s2 = if ($sentences.Count -ge 2) { $sentences[1] } else { "Пояснення панелей подано один раз на початку підрозділу без повторення службових фрагментів." }
        $s3 = if ($sentences.Count -ge 3) { $sentences[2] } else { "Результати використано для стислого порівняння комбінацій моделей у межах розділу." }

        $range = if ($firstCap.fig -eq $lastCap.fig) { "4.$($firstCap.fig)" } else { "4.$($firstCap.fig)–4.$($lastCap.fig)" }
        $s4 = "На рис. $range наведено результати для $($b.prefix)."
        if ($poolCites.Count -gt 0) {
          $s4 = $s4.TrimEnd(".") + " [" + ($poolCites -join ", ") + "]."
        }

        $final = (($s1.TrimEnd(".") + ".") + " " + ($s2.TrimEnd(".") + ".") + " " + ($s3.TrimEnd(".") + ".") + " " + $s4).Trim()
        $final = ($final -replace "\s{2,}", " ").Trim()

        $pSum = New-ParagraphWithStyleAndText -Xml $xml -Nsm $nsm -StyleVal $bodyStyle -Text $final
        [void]$body.InsertBefore($pSum, $paras[$insertBeforeIdx])
        $stats.inserted_summaries++

        # Refresh after insertion (indices shift).
        $paras = @($body.SelectNodes("./w:p", $nsm))
        $endIdx++
      }
    }

    # Delete boilerplate paragraphs in the subsection (style-safe).
    $paras = @($body.SelectNodes("./w:p", $nsm))
    $toDelete = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
    for ($i = $startIdx + 1; $i -le $endIdx; $i++) {
      $pt = Get-ParagraphText -Paragraph $paras[$i] -Nsm $nsm
      if (-not $pt) { continue }
      if ($pt -match "^4\.4\.[1-5]\b") { continue }
      if ($pt -match "^Рис\.") { continue }
      if (Paragraph-HasDrawing -Paragraph $paras[$i] -Nsm $nsm) { continue }
      if (Paragraph-HasBreak -Paragraph $paras[$i] -Nsm $nsm) { continue }
      if (Is-BoilerplateParagraphText -Text $pt) { $toDelete.Add($paras[$i]) | Out-Null }
    }
    foreach ($node in $toDelete) {
      [void]$node.ParentNode.RemoveChild($node)
      $stats.deleted_boilerplate_paras++
    }

    $stats.sections_processed += $secName
  }

  # Repo-wide citation safety pass (style-preserving): normalize ';' inside brackets + wrap bare lists when safe.
  $paras = @($body.SelectNodes("./w:p", $nsm))
  foreach ($p in $paras) {
    foreach ($tNode in @($p.SelectNodes(".//w:t", $nsm))) {
      $old = $tNode.InnerText
      if (-not $old) { continue }

      $new = $old
      $new = Normalize-BracketCitationsInText -Text @($new)
      if ($new -ne $old -and $old -match "\[.*;.*\]") { $stats.citation_semicolons_normalized++ }
      $new = Fix-BareCitationListsInString -Text $new -Stats ([ref]$stats)
      $new = Remove-RefTokensInText -Text @($new)

      if ($new -ne $old) { $tNode.InnerText = $new }
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
