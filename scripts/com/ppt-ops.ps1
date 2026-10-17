# ppt-ops.ps1 - PowerPoint inspect/edit operations (docs/FEATURES.md), dot-sourced by office.ps1.
# Keep this file ASCII (PowerShell 5.1 reads BOM-less scripts in the ANSI code page).
# Failures throw 'E|code|arg|...' (src/core/ppt.mjs turns them into Korean);
# edit wraps every failure as 'opfail|<n>|<op>|<message>' so the caller knows which operation failed.
# Every function name starts with Ppt: word/excel ops files share this script scope.

# spec layout key -> PpSlideLayout values that count as it (first: what Slides.Add creates when missing)
$PptLayoutKinds = [ordered]@{
  title = @(1, 1); title_content = @(16, 2); section = @(33, 33); two_content = @(29, 3)
  comparison = @(34, 34); title_only = @(11, 11); blank = @(12, 12)
}
$PptChartTypes = @{ column = 51; bar = 57; line = 4; pie = 5 }                    # XlChartType
$PptAutoShapes = @{ rectangle = 1; rounded_rectangle = 5; oval = 9; arrow = 33 }  # MsoAutoShapeType
$PptEffects = @{ none = 0; fade = 3849; push = 3855; wipe = 2817; split = 3587; cover = 1281 }  # PpEntryEffect
$PptAligns = @{ left = 1; center = 2; right = 3; justify = 4 }                    # PpParagraphAlignment
$PptSizes = @{ '16:9' = @(960, 540); '4:3' = @(720, 540) }
$script:PptLayoutCache = @{}

function Stop-Ppt { throw ('E|' + (@($args | ForEach-Object { [string]$_ }) -join '|')) }

function Get-PptSlide($pres, $n) {
  $count = $pres.Slides.Count
  $i = 0
  if (-not [int]::TryParse([string]$n, [ref]$i) -or $i -lt 1 -or $i -gt $count) { Stop-Ppt 'slide_range' $n $count }
  return $pres.Slides.Item($i)
}

function Get-PptTargets($pres, $ref) {
  if ([string]$ref -eq 'all') { return @($pres.Slides | ForEach-Object { $_ }) }
  return @(Get-PptSlide $pres $ref)
}

# shape by name (exact, then ignoring case) or by 1-based number
function Get-PptShape($slide, $ref) {
  $count = $slide.Shapes.Count
  if ($ref -is [int] -or $ref -is [long] -or $ref -is [double] -or $ref -is [decimal]) {
    $i = [int]$ref
    if ($i -lt 1 -or $i -gt $count) { Stop-Ppt 'shape_range' $slide.SlideIndex $i $count }
    return $slide.Shapes.Item($i)
  }
  $name = [string]$ref
  foreach ($sh in $slide.Shapes) { if ([string]$sh.Name -ceq $name) { return $sh } }
  foreach ($sh in $slide.Shapes) { if ([string]$sh.Name -ieq $name) { return $sh } }
  Stop-Ppt 'shape_not_found' $slide.SlideIndex $name
}

# points or 'NN%' of the slide width/height; -Positive: sizes must be > 0 (lines may go left/up)
function ConvertTo-PptPt($v, [double]$total, [string]$what, [switch]$Positive) {
  if ($null -eq $v) { return $null }
  $s = ([string]$v).Trim()
  $pct = $s.EndsWith('%')
  if ($pct) { $s = $s.TrimEnd('%').Trim() }
  $num = 0.0
  if (-not [double]::TryParse($s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) { Stop-Ppt 'bad_size' $what $v }
  if ($pct) { $num = $total * $num / 100.0 }
  if ($Positive -and $num -le 0) { Stop-Ppt 'size_not_positive' $what $v }
  return $num
}

# '#rrggbb' -> perceived brightness 0..1
function Get-PptLuma($c) {
  $h = ([string]$c).Trim().TrimStart('#')
  if ($h -notmatch '^[0-9a-fA-F]{6}$') { return 0.0 }
  $r = [Convert]::ToInt32($h.Substring(0, 2), 16); $g = [Convert]::ToInt32($h.Substring(2, 2), 16); $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
  return (0.299 * $r + 0.587 * $g + 0.114 * $b) / 255.0
}

# '#rrggbb' -> Office RGB (BGR order)
function ConvertTo-PptColor($c) {
  $h = ([string]$c).Trim().TrimStart('#')
  if ($h -notmatch '^[0-9a-fA-F]{6}$') { Stop-Ppt 'bad_color' $c }
  $r = [Convert]::ToInt32($h.Substring(0, 2), 16); $g = [Convert]::ToInt32($h.Substring(2, 2), 16); $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
  return [int]($r + $g * 256 + $b * 65536)
}

function ConvertTo-PptText($t) { if ($null -eq $t) { return '' }; return ([string]$t) -replace "`r`n|`n", "`r" }

function Get-PptShortText($t, [int]$max) {
  $s = ([string]$t) -replace "`r`n|`r|`v", "`n"
  if ($s.Length -gt $max) { $s = $s.Substring(0, $max) + '...' }
  return $s
}

function Set-PptFont($font, $op) {
  if ($null -ne $op.bold) { $font.Bold = $(if ($op.bold) { -1 } else { 0 }) }
  if ($null -ne $op.italic) { $font.Italic = $(if ($op.italic) { -1 } else { 0 }) }
  if ($null -ne $op.underline) { $font.Underline = $(if ($op.underline) { -1 } else { 0 }) }
  if ($op.size) { $font.Size = [double]$op.size }
  if ($op.color) { $font.Color.RGB = ConvertTo-PptColor $op.color }
  if ($op.font) { $font.Name = [string]$op.font; try { $font.NameFarEast = [string]$op.font } catch {} }
}

function Get-PptPlaceholder($slide, [int[]]$types) {
  foreach ($sh in $slide.Shapes) {
    if ($sh.Type -eq 14) {
      $t = 0; try { $t = [int]$sh.PlaceholderFormat.Type } catch {}
      if ($types -contains $t) { return $sh }
    }
  }
  return $null
}

# content placeholders in fill order: objects, then bodies, then subtitle
function Get-PptBodyShapes($slide) {
  $list = New-Object System.Collections.ArrayList
  foreach ($rank in @((7, 17), (2, 6), (4, 4))) {
    foreach ($sh in $slide.Shapes) {
      if ($sh.Type -ne 14 -or -not $sh.HasTextFrame) { continue }
      $t = 0; try { $t = [int]$sh.PlaceholderFormat.Type } catch {}
      if ($rank -contains $t) { [void]$list.Add($sh) }
    }
  }
  return ,$list
}

# body: string (one paragraph per line) or array of strings / {text, level}
function Get-PptBulletItems($v) {
  $items = New-Object System.Collections.ArrayList
  if ($null -eq $v) { return ,$items }
  if ($v -is [string]) {
    foreach ($line in ($v -split "`r`n|`n|`r")) { [void]$items.Add(@{ text = $line; level = 1 }) }
    return ,$items
  }
  foreach ($b in @($v)) {
    if ($null -eq $b) { continue }
    if ($b -is [string]) { [void]$items.Add(@{ text = $b; level = 1 }); continue }
    $lv = 1; if ($b.level) { $lv = [Math]::Max(1, [Math]::Min(9, [int]$b.level)) }
    [void]$items.Add(@{ text = [string]$b.text; level = $lv })
  }
  return ,$items
}

function Set-PptBullets($tr, $v) {
  $items = Get-PptBulletItems $v
  $vt = [string][char]11
  $tr.Text = (@($items | ForEach-Object { ([string]$_.text) -replace "`r`n|`n|`r", $vt }) -join "`r")
  for ($i = 0; $i -lt $items.Count; $i++) { $tr.Paragraphs($i + 1, 1).IndentLevel = $items[$i].level }
  return $items.Count
}

function Set-PptTitle($slide, $text) {
  $sh = Get-PptPlaceholder $slide @(1, 3, 5)
  if (-not $sh) { try { $sh = $slide.Shapes.AddTitle() } catch { $sh = $null } }
  if (-not $sh) { Stop-Ppt 'no_title' $slide.SlideIndex }
  $sh.TextFrame.TextRange.Text = ConvertTo-PptText $text
  return $sh
}

function Set-PptBody($slide, $value, $n) {
  $cands = Get-PptBodyShapes $slide
  $k = 1; if ($n) { $k = [int]$n }
  if ($cands.Count -lt $k) { Stop-Ppt 'no_body' $slide.SlideIndex $k $cands.Count }
  $sh = $cands[$k - 1]
  $lines = Set-PptBullets $sh.TextFrame.TextRange $value
  return @{ shape = [string]$sh.Name; lines = [int]$lines }
}

function Get-PptNotesShape($slide) {
  foreach ($ph in $slide.NotesPage.Shapes.Placeholders) {
    $t = 0; try { $t = [int]$ph.PlaceholderFormat.Type } catch {}
    if ($t -eq 2) { return $ph }
  }
  return $null
}

function Set-PptNotes($slide, $text) {
  $ph = Get-PptNotesShape $slide
  if (-not $ph) { Stop-Ppt 'no_notes' $slide.SlideIndex }
  $ph.TextFrame.TextRange.Text = ConvertTo-PptText $text
}

function Get-PptNotes($slide) {
  try { $ph = Get-PptNotesShape $slide; if ($ph -and $ph.TextFrame.HasText) { return [string]$ph.TextFrame.TextRange.Text } } catch {}
  return ''
}

# ---------------------------------------------------------------- layouts
# PpSlideLayout of every custom layout: CustomLayout has no type property, so add a scratch slide,
# read Slide.Layout and delete it again (works on read-only and windowless presentations).
function Get-PptLayoutTypes($pres, $design) {
  $key = [string]$design.Index
  if ($script:PptLayoutCache.ContainsKey($key)) { return ,$script:PptLayoutCache[$key] }
  $cls = $design.SlideMaster.CustomLayouts
  $types = New-Object 'int[]' $cls.Count
  for ($i = 1; $i -le $cls.Count; $i++) {
    $tmp = $null
    try { $tmp = $pres.Slides.AddSlide($pres.Slides.Count + 1, $cls.Item($i)); $types[$i - 1] = [int]$tmp.Layout }
    catch { $types[$i - 1] = 0 }
    finally { if ($tmp) { try { $tmp.Delete() } catch {} } }
  }
  $script:PptLayoutCache[$key] = $types
  return ,$types
}

# user-made layouts (type 'custom'): guess the key from their placeholders
function Get-PptLayoutShapeKey($cl) {
  $title = 0; $ctr = 0; $sub = 0; $obj = 0; $body = 0
  foreach ($sh in $cl.Shapes) {
    if ($sh.Type -ne 14) { continue }
    $t = 0; try { $t = [int]$sh.PlaceholderFormat.Type } catch {}
    switch ($t) {
      { $_ -in @(1, 5) } { $title++ }
      3 { $ctr++ }
      4 { $sub++ }
      { $_ -in @(7, 8, 12, 17, 18) } { $obj++ }
      { $_ -in @(2, 6) } { $body++ }
    }
  }
  if ($ctr -ge 1 -and $sub -ge 1) { return 'title' }
  if ($title + $ctr -eq 0 -and $obj + $body + $sub -eq 0) { return 'blank' }
  if ($title -ge 1 -and $obj + $body + $sub -eq 0) { return 'title_only' }
  if ($title -ge 1 -and $obj -eq 1 -and $body -eq 0) { return 'title_content' }
  if ($title -ge 1 -and $obj -eq 2 -and $body -eq 0) { return 'two_content' }
  if ($title -ge 1 -and $obj -eq 2 -and $body -eq 2) { return 'comparison' }
  return ''
}

function Get-PptLayoutKey([int]$type, $cl) {
  foreach ($k in $PptLayoutKinds.Keys) { if ($PptLayoutKinds[$k] -contains $type) { return $k } }
  if ($type -eq 32) { $g = Get-PptLayoutShapeKey $cl; if ($g) { return $g } }
  return ''
}

# layout key (by layout type, then placeholder shape) or exact layout name in one design
function Find-PptLayout($pres, $design, [string]$key) {
  $cls = $design.SlideMaster.CustomLayouts
  if ($PptLayoutKinds.Contains($key)) {
    $types = Get-PptLayoutTypes $pres $design
    foreach ($want in $PptLayoutKinds[$key]) {
      for ($i = 0; $i -lt $types.Count; $i++) { if ($types[$i] -eq $want) { return $cls.Item($i + 1) } }
    }
    for ($i = 0; $i -lt $types.Count; $i++) {
      if ($types[$i] -eq 32 -and (Get-PptLayoutShapeKey $cls.Item($i + 1)) -eq $key) { return $cls.Item($i + 1) }
    }
    return $null
  }
  for ($i = 1; $i -le $cls.Count; $i++) { if ([string]$cls.Item($i).Name -ceq $key) { return $cls.Item($i) } }
  for ($i = 1; $i -le $cls.Count; $i++) { if ([string]$cls.Item($i).Name -ieq $key) { return $cls.Item($i) } }
  return $null
}

function Get-PptLayoutNames($pres) {
  $names = New-Object System.Collections.ArrayList
  foreach ($d in $pres.Designs) { foreach ($cl in $d.SlideMaster.CustomLayouts) { if (-not $names.Contains([string]$cl.Name)) { [void]$names.Add([string]$cl.Name) } } }
  return ($names -join ', ')
}

# ---------------------------------------------------------------- Excel started for chart data
function Get-PptExcelIds { return @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }

# chart data runs in an Excel that PowerPoint itself starts ('EXCEL.EXE /Automation -Embedding /K', child of POWERPNT)
function Get-PptChartExcelIds($before) {
  $ppt = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
  $kids = @()
  try {
    $kids = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction Stop |
      Where-Object { $ppt -contains [int]$_.ParentProcessId } | ForEach-Object { [int]$_.ProcessId })
  } catch {
    $kids = @(Get-PptExcelIds)
  }
  return @($kids | Where-Object { $before -notcontains $_ })
}

# That Excel keeps running after Quit and keeps PowerPoint alive: end it (only processes that did not exist before).
function Stop-PptExcel($ids) {
  $ids = @($ids)
  if (-not $ids.Count) { return }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  foreach ($id in $ids) {
    try {
      $p = Get-Process -Id $id -ErrorAction Stop
      if (-not $p.WaitForExit(1500)) { $p.Kill(); [void]$p.WaitForExit(5000) }
    } catch {}
  }
}

# ---------------------------------------------------------------- edit operations
function Invoke-PptAddSlide($pres, $op) {
  $count = $pres.Slides.Count
  $pos = $count + 1
  if ($null -ne $op.position) {
    $pos = [int]$op.position
    if ($pos -lt 1 -or $pos -gt $count + 1) { Stop-Ppt 'position_range' $op.position ($count + 1) }
  }
  $key = 'title_content'; if ($op.layout) { $key = [string]$op.layout }
  # same design as the slide before the new one
  if ($count -gt 0) { $design = $pres.Slides.Item([Math]::Max(1, [Math]::Min($pos - 1, $count))).Design }
  else { $design = $pres.Designs.Item(1) }
  $cl = Find-PptLayout $pres $design $key
  if (-not $cl -and -not $PptLayoutKinds.Contains($key)) {
    foreach ($d in $pres.Designs) { if (-not $cl) { $cl = Find-PptLayout $pres $d $key } }
  }
  $created = $false
  if ($cl) { $slide = $pres.Slides.AddSlide($pos, $cl) }
  elseif ($PptLayoutKinds.Contains($key)) {
    # the master has no such layout: the legacy Slides.Add makes PowerPoint add a standard one
    $slide = $pres.Slides.Add($pos, $PptLayoutKinds[$key][0])
    $created = $true
    $script:PptLayoutCache = @{}
  } else { Stop-Ppt 'layout_not_found' $key (Get-PptLayoutNames $pres) }
  if ($null -ne $op.title) { [void](Set-PptTitle $slide $op.title) }
  if ($null -ne $op.body) { [void](Set-PptBody $slide $op.body 1) }
  if ($null -ne $op.notes) { Set-PptNotes $slide $op.notes }
  return @{ op = 'add_slide'; slide = [int]$slide.SlideIndex; layout = [string]$slide.CustomLayout.Name; created = $created; count = [int]$pres.Slides.Count }
}

function Invoke-PptAddTextbox($pres, $op) {
  $slide = Get-PptSlide $pres $op.slide
  $ps = $pres.PageSetup
  $l = ConvertTo-PptPt $op.left $ps.SlideWidth 'left'; $t = ConvertTo-PptPt $op.top $ps.SlideHeight 'top'
  $w = ConvertTo-PptPt $op.width $ps.SlideWidth 'width' -Positive; $h = ConvertTo-PptPt $op.height $ps.SlideHeight 'height' -Positive
  $sh = $slide.Shapes.AddTextbox(1, $l, $t, $w, $h)
  $sh.TextFrame.WordWrap = -1
  $tr = $sh.TextFrame.TextRange
  $tr.Text = ConvertTo-PptText $op.text
  Set-PptFont $tr.Font $op
  if ($op.align) { $tr.ParagraphFormat.Alignment = $PptAligns[[string]$op.align] }
  # AddTextbox shrinks the box to the text (autosize): keep the asked height unless the text needs more
  if ([double]$sh.Height -lt $h) { $sh.TextFrame.AutoSize = 0; $sh.Height = $h }
  $grown = ([double]$sh.Height -gt $h + 0.5)
  return @{ op = 'add_textbox'; slide = [int]$slide.SlideIndex; shape = [string]$sh.Name; left = $l; top = $t; width = [double]$sh.Width; height = [double]$sh.Height; asked = $h; grown = $grown }
}

function Invoke-PptAddImage($pres, $op) {
  $slide = Get-PptSlide $pres $op.slide
  $file = [string]$op.image_path
  if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { Stop-Ppt 'no_file' $file }
  $ps = $pres.PageSetup
  $l = ConvertTo-PptPt $op.left $ps.SlideWidth 'left'; $t = ConvertTo-PptPt $op.top $ps.SlideHeight 'top'
  $w = ConvertTo-PptPt $op.width $ps.SlideWidth 'width' -Positive; $h = ConvertTo-PptPt $op.height $ps.SlideHeight 'height' -Positive
  # FileName, LinkToFile(no), SaveWithDocument(yes), Left, Top, Width/Height -1 = picture size
  try { $pic = $slide.Shapes.AddPicture($file, 0, -1, $l, $t, -1, -1) } catch { Stop-Ppt 'image_failed' $_.Exception.Message }
  if ($null -ne $w -and $null -ne $h) { $pic.LockAspectRatio = 0; $pic.Width = $w; $pic.Height = $h }
  elseif ($null -ne $w) { $pic.LockAspectRatio = -1; $pic.Width = $w }
  elseif ($null -ne $h) { $pic.LockAspectRatio = -1; $pic.Height = $h }
  else {
    # no size given: shrink (keeping the ratio) when the picture runs off the slide
    $maxW = $ps.SlideWidth - $l; $maxH = $ps.SlideHeight - $t
    if ($maxW -gt 0 -and $maxH -gt 0 -and ($pic.Width -gt $maxW -or $pic.Height -gt $maxH)) {
      $pic.LockAspectRatio = -1
      $pic.Width = $pic.Width * [Math]::Min($maxW / $pic.Width, $maxH / $pic.Height)
    }
  }
  return @{ op = 'add_image'; slide = [int]$slide.SlideIndex; shape = [string]$pic.Name; left = $l; top = $t; width = [double]$pic.Width; height = [double]$pic.Height }
}

function Invoke-PptAddTable($pres, $op) {
  $slide = Get-PptSlide $pres $op.slide
  $ps = $pres.PageSetup
  $l = ConvertTo-PptPt $op.left $ps.SlideWidth 'left'; $t = ConvertTo-PptPt $op.top $ps.SlideHeight 'top'
  $w = ConvertTo-PptPt $op.width $ps.SlideWidth 'width' -Positive; $h = ConvertTo-PptPt $op.height $ps.SlideHeight 'height' -Positive
  $rows = @($op.rows)
  $nr = $rows.Count; $nc = 1
  foreach ($r in $rows) { $nc = [Math]::Max($nc, @($r).Count) }
  if ($null -ne $h) { $sh = $slide.Shapes.AddTable($nr, $nc, $l, $t, $w, $h) } else { $sh = $slide.Shapes.AddTable($nr, $nc, $l, $t, $w) }
  $tbl = $sh.Table
  for ($r = 0; $r -lt $nr; $r++) {
    $cells = @($rows[$r])
    for ($c = 0; $c -lt $cells.Count; $c++) { $tbl.Cell($r + 1, $c + 1).Shape.TextFrame.TextRange.Text = ConvertTo-PptText $cells[$c] }
  }
  if ($op.header_bold -ne $false) { for ($c = 1; $c -le $nc; $c++) { $tbl.Cell(1, $c).Shape.TextFrame.TextRange.Font.Bold = -1 } }
  return @{ op = 'add_table'; slide = [int]$slide.SlideIndex; shape = [string]$sh.Name; rows = $nr; cols = $nc }
}

function Invoke-PptTableSetCell($pres, $op) {
  $slide = Get-PptSlide $pres $op.slide
  $sh = Get-PptShape $slide $op.shape
  if (-not $sh.HasTable) { Stop-Ppt 'not_table' $slide.SlideIndex $sh.Name }
  $tbl = $sh.Table
  $r = [int]$op.row; $c = [int]$op.col
  if ($r -lt 1 -or $r -gt $tbl.Rows.Count -or $c -lt 1 -or $c -gt $tbl.Columns.Count) { Stop-Ppt 'cell_range' $r $c $tbl.Rows.Count $tbl.Columns.Count }
  $tbl.Cell($r, $c).Shape.TextFrame.TextRange.Text = ConvertTo-PptText $op.text
  return @{ op = 'table_set_cell'; slide = [int]$slide.SlideIndex; shape = [string]$sh.Name; row = $r; col = $c }
}

function Invoke-PptAddShape($pres, $op) {
  $slide = Get-PptSlide $pres $op.slide
  $ps = $pres.PageSetup
  $l = ConvertTo-PptPt $op.left $ps.SlideWidth 'left'; $t = ConvertTo-PptPt $op.top $ps.SlideHeight 'top'
  $kind = [string]$op.shape
  if ($kind -eq 'line') {
    $w = ConvertTo-PptPt $op.width $ps.SlideWidth 'width'; $h = ConvertTo-PptPt $op.height $ps.SlideHeight 'height'
    if ($null -ne $op.text -and [string]$op.text -ne '') { Stop-Ppt 'line_text' }
    $sh = $slide.Shapes.AddLine($l, $t, $l + $w, $t + $h)
  } else {
    $w = ConvertTo-PptPt $op.width $ps.SlideWidth 'width' -Positive; $h = ConvertTo-PptPt $op.height $ps.SlideHeight 'height' -Positive
    if (-not $PptAutoShapes.ContainsKey($kind)) { Stop-Ppt 'bad_shape' $kind }
    $sh = $slide.Shapes.AddShape($PptAutoShapes[$kind], $l, $t, $w, $h)
  }
  if ($op.fill -and $kind -ne 'line') {
    if ([string]$op.fill -eq 'none') { $sh.Fill.Visible = 0 }
    else { $sh.Fill.Visible = -1; $sh.Fill.Solid(); $sh.Fill.ForeColor.RGB = ConvertTo-PptColor $op.fill; $sh.Fill.Transparency = 0 }
  }
  if ($op.line) {
    if ([string]$op.line -eq 'none') { $sh.Line.Visible = 0 }
    else { $sh.Line.Visible = -1; $sh.Line.ForeColor.RGB = ConvertTo-PptColor $op.line }
  }
  if ($null -ne $op.text -and $kind -ne 'line') {
    $tr = $sh.TextFrame.TextRange
    $tr.Text = ConvertTo-PptText $op.text
    # the default shape style writes white text: make it readable without a fill or on a light fill
    if ($op.color) { $tr.Font.Color.RGB = ConvertTo-PptColor $op.color }
    elseif ([string]$op.fill -eq 'none') { $tr.Font.Color.ObjectThemeColor = 13 }   # msoThemeColorText1 (follows the background)
    elseif ($op.fill -and (Get-PptLuma $op.fill) -gt 0.6) { $tr.Font.Color.ObjectThemeColor = 1 }   # msoThemeColorDark1
  }
  return @{ op = 'add_shape'; slide = [int]$slide.SlideIndex; shape = [string]$sh.Name }
}

function Invoke-PptAddChart($pres, $op) {
  $slide = Get-PptSlide $pres $op.slide
  $ps = $pres.PageSetup
  $l = ConvertTo-PptPt $op.left $ps.SlideWidth 'left'; $t = ConvertTo-PptPt $op.top $ps.SlideHeight 'top'
  $w = ConvertTo-PptPt $op.width $ps.SlideWidth 'width' -Positive; $h = ConvertTo-PptPt $op.height $ps.SlideHeight 'height' -Positive
  $type = $PptChartTypes[[string]$op.type]
  if (-not $type) { Stop-Ppt 'bad_chart' $op.type }
  $cats = @($op.categories); $series = @($op.series)
  # AddChart2 fails (E_FAIL) without a document window: open one minimized; it closes with the presentation
  if ($pres.Windows.Count -eq 0) { $win = $pres.NewWindow(); try { $win.WindowState = 2 } catch {} }
  $before = Get-PptExcelIds
  $sh = $null; $wb = $null; $ws = $null; $xl = $null; $started = $null
  try {
    try { $sh = $slide.Shapes.AddChart2(-1, $type, $l, $t, $w, $h, $true) } catch { Stop-Ppt 'chart_failed' $_.Exception.Message }
    $started = @(Get-PptChartExcelIds $before)
    $chart = $sh.Chart
    $wb = $chart.ChartData.Workbook
    $xl = $wb.Application
    if ($started.Count) { try { $xl.Visible = $false } catch {} }
    $ws = $wb.Worksheets.Item(1)
    # sample data sits in a table: make it a plain range and clear it
    while ($ws.ListObjects.Count -gt 0) { $ws.ListObjects.Item(1).Unlist() }
    [void]$ws.Cells.Clear()
    $nr = $cats.Count + 1; $nc = $series.Count + 1
    $data = New-Object 'object[,]' $nr, $nc
    for ($c = 0; $c -lt $series.Count; $c++) {
      $data[0, ($c + 1)] = [string]$series[$c].name
      $vals = @($series[$c].values)
      for ($r = 0; $r -lt $cats.Count; $r++) { if ($r -lt $vals.Count -and $null -ne $vals[$r]) { $data[($r + 1), ($c + 1)] = [double]$vals[$r] } }
    }
    for ($r = 0; $r -lt $cats.Count; $r++) { $data[($r + 1), 0] = [string]$cats[$r] }
    # category labels and series names stay text (a numeric year column would turn into a series)
    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($nr, 1)).NumberFormat = '@'
    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $nc)).NumberFormat = '@'
    $rng = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($nr, $nc))
    $rng.Value2 = $data
    # series in columns (xlColumns); check it took (the sample has 3 series) and retry once in a while
    $src = "='" + ([string]$ws.Name).Replace("'", "''") + "'!" + $rng.Address()
    $want = $series.Count
    $chart.SetSourceData($src, 2)
    for ($try = 0; $try -lt 3 -and [int]$chart.SeriesCollection().Count -ne $want; $try++) { Start-Sleep -Milliseconds 300; $chart.SetSourceData($src, 2) }
    while ([int]$chart.SeriesCollection().Count -gt $want) { $chart.SeriesCollection([int]$chart.SeriesCollection().Count).Delete() }
    $got = [int]$chart.SeriesCollection().Count
    if ($got -ne $want) { Stop-Ppt 'chart_failed' "series $got/$want" }
    if ($op.title) { $chart.HasTitle = $true; $chart.ChartTitle.Text = [string]$op.title } else { $chart.HasTitle = $false }
  } finally {
    if ($wb) { try { $wb.Close() } catch {} }
    if ($null -eq $started) { $started = @(Get-PptChartExcelIds $before) }
    if ($xl -and $started.Count) { try { if ($xl.Workbooks.Count -eq 0) { $xl.Quit() } } catch {} }
    Release $ws; Release $wb; Release $xl
    $ws = $null; $wb = $null; $xl = $null
    Stop-PptExcel $started
  }
  return @{ op = 'add_chart'; slide = [int]$slide.SlideIndex; shape = [string]$sh.Name; series = $got; categories = $cats.Count }
}

function Add-PptTextRanges($sh, $list) {
  if ($sh.Type -eq 6) { foreach ($it in $sh.GroupItems) { Add-PptTextRanges $it $list }; return }
  if ($sh.HasTable) {
    foreach ($row in $sh.Table.Rows) { foreach ($cell in $row.Cells) { [void]$list.Add($cell.Shape.TextFrame.TextRange) } }
    return
  }
  if ($sh.HasTextFrame) { if ($sh.TextFrame.HasText) { [void]$list.Add($sh.TextFrame.TextRange) } }
}

function Invoke-PptFormatText($pres, $op) {
  $find = [string]$op.find
  if ($null -ne $op.slide) { $slides = @(Get-PptSlide $pres $op.slide) } else { $slides = @($pres.Slides | ForEach-Object { $_ }) }
  $n = 0
  $hitSlides = New-Object System.Collections.ArrayList
  :outer foreach ($s in $slides) {
    $ranges = New-Object System.Collections.ArrayList
    foreach ($sh in $s.Shapes) { Add-PptTextRanges $sh $ranges }
    foreach ($tr in $ranges) {
      $after = 0
      while ($true) {
        # FindWhat, After, MatchCase(msoTrue), WholeWords(msoFalse)
        $hit = $tr.Find($find, $after, -1, 0)
        if (-not $hit) { break }
        Set-PptFont $hit.Font $op
        $n++
        if (-not $hitSlides.Contains([int]$s.SlideIndex)) { [void]$hitSlides.Add([int]$s.SlideIndex) }
        if (-not $op.all) { break outer }
        $after = $hit.Start + $hit.Length - 1
      }
    }
  }
  if ($n -eq 0) { Stop-Ppt 'not_found' $find }
  return @{ op = 'format_text'; count = $n; slides = @($hitSlides) }
}

function Invoke-PptSetBackground($pres, $op) {
  $rgb = ConvertTo-PptColor $op.color
  $targets = @(Get-PptTargets $pres $op.slide)
  foreach ($s in $targets) {
    $s.FollowMasterBackground = 0
    $f = $s.Background.Fill
    $f.Visible = -1; $f.Solid(); $f.ForeColor.RGB = $rgb; $f.Transparency = 0
  }
  if ([string]$op.slide -eq 'all') {
    # masters too, so slides added later match
    foreach ($d in $pres.Designs) { try { $f = $d.SlideMaster.Background.Fill; $f.Solid(); $f.ForeColor.RGB = $rgb } catch {} }
  }
  return @{ op = 'set_background'; slides = @($targets | ForEach-Object { [int]$_.SlideIndex }); all = ([string]$op.slide -eq 'all') }
}

function Invoke-PptSetTransition($pres, $op) {
  $eff = $PptEffects[[string]$op.effect]
  if ($null -eq $eff) { Stop-Ppt 'bad_effect' $op.effect }
  $targets = @(Get-PptTargets $pres $op.slide)
  foreach ($s in $targets) {
    $tr = $s.SlideShowTransition
    $tr.EntryEffect = $eff
    if ($null -ne $op.duration_s -and $eff -ne 0) { $tr.Duration = [double]$op.duration_s }
  }
  return @{ op = 'set_transition'; slides = @($targets | ForEach-Object { [int]$_.SlideIndex }); all = ([string]$op.slide -eq 'all') }
}

function Invoke-PptOp($pres, $op) {
  switch ([string]$op.op) {
    'add_slide' { return Invoke-PptAddSlide $pres $op }
    'delete_slide' {
      $s = Get-PptSlide $pres $op.slide
      $n = [int]$s.SlideIndex
      $s.Delete()
      return @{ op = 'delete_slide'; slide = $n; count = [int]$pres.Slides.Count }
    }
    'duplicate_slide' {
      $s = Get-PptSlide $pres $op.slide
      $from = [int]$s.SlideIndex
      $dup = $s.Duplicate().Item(1)
      if ($null -ne $op.position) {
        $to = [int]$op.position
        if ($to -lt 1 -or $to -gt $pres.Slides.Count) { Stop-Ppt 'position_range' $op.position $pres.Slides.Count }
        $dup.MoveTo($to)
      }
      return @{ op = 'duplicate_slide'; from = $from; slide = [int]$dup.SlideIndex; count = [int]$pres.Slides.Count }
    }
    'move_slide' {
      $s = Get-PptSlide $pres $op.slide
      $from = [int]$s.SlideIndex
      $to = [int]$op.to
      if ($to -lt 1 -or $to -gt $pres.Slides.Count) { Stop-Ppt 'position_range' $op.to $pres.Slides.Count }
      $s.MoveTo($to)
      return @{ op = 'move_slide'; from = $from; slide = [int]$s.SlideIndex }
    }
    'set_title' {
      $s = Get-PptSlide $pres $op.slide
      $sh = Set-PptTitle $s $op.text
      return @{ op = 'set_title'; slide = [int]$s.SlideIndex; shape = [string]$sh.Name }
    }
    'set_body' {
      $s = Get-PptSlide $pres $op.slide
      $b = Set-PptBody $s $op.bullets $op.placeholder
      return @{ op = 'set_body'; slide = [int]$s.SlideIndex; shape = $b.shape; lines = $b.lines }
    }
    'set_notes' {
      $s = Get-PptSlide $pres $op.slide
      Set-PptNotes $s $op.text
      return @{ op = 'set_notes'; slide = [int]$s.SlideIndex }
    }
    'set_text' {
      $s = Get-PptSlide $pres $op.slide
      $sh = Get-PptShape $s $op.shape
      if ($sh.HasTable) { Stop-Ppt 'is_table' $s.SlideIndex $sh.Name }
      if (-not $sh.HasTextFrame) { Stop-Ppt 'no_text' $s.SlideIndex $sh.Name }
      $sh.TextFrame.TextRange.Text = ConvertTo-PptText $op.text
      return @{ op = 'set_text'; slide = [int]$s.SlideIndex; shape = [string]$sh.Name }
    }
    'add_textbox' { return Invoke-PptAddTextbox $pres $op }
    'add_image' { return Invoke-PptAddImage $pres $op }
    'add_table' { return Invoke-PptAddTable $pres $op }
    'table_set_cell' { return Invoke-PptTableSetCell $pres $op }
    'add_shape' { return Invoke-PptAddShape $pres $op }
    'add_chart' { return Invoke-PptAddChart $pres $op }
    'format_text' { return Invoke-PptFormatText $pres $op }
    'delete_shape' {
      $s = Get-PptSlide $pres $op.slide
      $sh = Get-PptShape $s $op.shape
      $name = [string]$sh.Name
      $sh.Delete()
      return @{ op = 'delete_shape'; slide = [int]$s.SlideIndex; shape = $name }
    }
    'set_background' { return Invoke-PptSetBackground $pres $op }
    'set_transition' { return Invoke-PptSetTransition $pres $op }
    'hide_slide' {
      $s = Get-PptSlide $pres $op.slide
      $hide = ($op.hidden -ne $false)
      $s.SlideShowTransition.Hidden = $(if ($hide) { -1 } else { 0 })
      return @{ op = 'hide_slide'; slide = [int]$s.SlideIndex; hidden = $hide }
    }
    'apply_theme' {
      $file = [string]$op.theme_path
      if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { Stop-Ppt 'no_file' $file }
      try { $pres.ApplyTheme($file) } catch {
        $first = $_.Exception.Message
        try { $pres.ApplyTemplate($file) } catch { Stop-Ppt 'theme_failed' $first }
      }
      $script:PptLayoutCache = @{}
      return @{ op = 'apply_theme'; design = [string]$pres.Designs.Item(1).Name; layouts = [int]$pres.Designs.Item(1).SlideMaster.CustomLayouts.Count }
    }
    'set_slide_size' {
      $size = $PptSizes[[string]$op.preset]
      if (-not $size) { Stop-Ppt 'bad_preset' $op.preset }
      $oldW = [double]$pres.PageSetup.SlideWidth; $oldH = [double]$pres.PageSetup.SlideHeight
      # narrower: PowerPoint scales content down; wider: content keeps its size and is centred
      $pres.PageSetup.SlideWidth = $size[0]
      $pres.PageSetup.SlideHeight = $size[1]
      return @{ op = 'set_slide_size'; width = [double]$pres.PageSetup.SlideWidth; height = [double]$pres.PageSetup.SlideHeight; oldWidth = $oldW; oldHeight = $oldH }
    }
    default { Stop-Ppt 'unknown_op' $op.op }
  }
}

# Runs every operation in order; any failure throws before office.ps1 saves.
function Invoke-PptEdit($pres, $app, $req) {
  $script:PptLayoutCache = @{}
  $results = New-Object System.Collections.ArrayList
  $i = 0
  foreach ($op in @($req.operations)) {
    $i++
    try {
      # keep only the result table even if a COM call leaked a value into the pipeline
      $r = @(Invoke-PptOp $pres $op) | Where-Object { $_ -is [hashtable] } | Select-Object -Last 1
      [void]$results.Add($r)
    } catch {
      throw "opfail|$i|$($op.op)|$($_.Exception.Message)"
    }
  }
  return $results.ToArray()
}

# ---------------------------------------------------------------- inspect
function Get-PptShapeInfo($sh, [int]$k) {
  $info = @{ index = $k; name = [string]$sh.Name; type = [int]$sh.Type;
    left = [Math]::Round([double]$sh.Left, 1); top = [Math]::Round([double]$sh.Top, 1);
    width = [Math]::Round([double]$sh.Width, 1); height = [Math]::Round([double]$sh.Height, 1) }
  if ($sh.Type -eq 14) { try { $info.placeholder = [int]$sh.PlaceholderFormat.Type } catch {} }
  if ($sh.Type -eq 6) {
    $texts = @()
    foreach ($it in $sh.GroupItems) { try { if ($it.HasTextFrame -and $it.TextFrame.HasText) { $texts += [string]$it.TextFrame.TextRange.Text } } catch {} }
    $info.items = [int]$sh.GroupItems.Count
    if ($texts.Count) { $info.text = Get-PptShortText ($texts -join ' / ') 200 }
  } elseif ($sh.HasTable) {
    $tbl = $sh.Table
    $info.rows = [int]$tbl.Rows.Count; $info.cols = [int]$tbl.Columns.Count
    $first = @()
    for ($c = 1; $c -le $tbl.Columns.Count; $c++) { $first += [string]$tbl.Cell(1, $c).Shape.TextFrame.TextRange.Text }
    $info.text = Get-PptShortText ($first -join ' | ') 200
  } elseif ($sh.HasChart) {
    $ch = $sh.Chart
    $info.chart = [int]$ch.ChartType
    try { if ($ch.HasTitle) { $info.text = Get-PptShortText $ch.ChartTitle.Text 200 } } catch {}
  } elseif ($sh.HasTextFrame) {
    if ($sh.TextFrame.HasText) { $info.text = Get-PptShortText $sh.TextFrame.TextRange.Text 200 }
  }
  return $info
}

function Invoke-PptInspect($pres, $app, $req) {
  $script:PptLayoutCache = @{}
  $ps = $pres.PageSetup
  $designs = New-Object System.Collections.ArrayList
  $layouts = New-Object System.Collections.ArrayList
  foreach ($d in $pres.Designs) {
    [void]$designs.Add([string]$d.Name)
    $types = Get-PptLayoutTypes $pres $d
    $cls = $d.SlideMaster.CustomLayouts
    for ($i = 1; $i -le $cls.Count; $i++) {
      $cl = $cls.Item($i)
      [void]$layouts.Add(@{ design = [string]$d.Name; index = $i; name = [string]$cl.Name; type = [int]$types[$i - 1]; key = (Get-PptLayoutKey $types[$i - 1] $cl) })
    }
  }
  $count = $pres.Slides.Count
  # one slide, or slides offset+1 .. offset+limit (paging keeps the reply short on big decks)
  $from = 1; $to = $count
  if ($req.slide) { $from = [int](Get-PptSlide $pres $req.slide).SlideIndex; $to = $from }
  else {
    if ($req.offset) { $from = [int]$req.offset + 1 }
    if ($req.limit) { $to = [Math]::Min($count, $from + [int]$req.limit - 1) }
  }
  $maxShapes = 0; if ($req.maxShapes) { $maxShapes = [int]$req.maxShapes }
  $slides = New-Object System.Collections.ArrayList
  for ($n = $from; $n -le $to; $n++) {
    $s = $pres.Slides.Item($n)
    $title = ''
    $tsh = Get-PptPlaceholder $s @(1, 3, 5)
    if ($tsh -and $tsh.HasTextFrame -and $tsh.TextFrame.HasText) { $title = Get-PptShortText $tsh.TextFrame.TextRange.Text 200 }
    $sst = $s.SlideShowTransition
    $duration = 0.0; try { $duration = [Math]::Round([double]$sst.Duration, 2) } catch {}
    $shapes = New-Object System.Collections.ArrayList
    $total = [int]$s.Shapes.Count
    $show = $total; if ($maxShapes -gt 0) { $show = [Math]::Min($total, $maxShapes) }
    for ($k = 1; $k -le $show; $k++) { [void]$shapes.Add((Get-PptShapeInfo $s.Shapes.Item($k) $k)) }
    [void]$slides.Add(@{ index = $n; layout = [string]$s.CustomLayout.Name; layoutType = [int]$s.Layout;
      key = (Get-PptLayoutKey ([int]$s.Layout) $s.CustomLayout); design = [string]$s.Design.Name; title = $title;
      hidden = ([int]$sst.Hidden -eq -1); effect = [int]$sst.EntryEffect; duration = $duration;
      notes = (Get-PptShortText (Get-PptNotes $s) 500); shapeCount = $total; shapes = $shapes })
  }
  return @{ slideWidth = [double]$ps.SlideWidth; slideHeight = [double]$ps.SlideHeight; slideCount = $count;
    designs = $designs; layouts = $layouts; slides = $slides }
}
