# excel-ops.ps1 - Excel-specific inspect/edit operations, dot-sourced by office.ps1.
# Keep this file ASCII (PowerShell 5.1 reads BOM-less scripts in the ANSI code page).
# Operations arrive already checked and normalized by src/core/excel.mjs (colors 'RRGGBB', ranges 'A1:B2',
# columns 'A:C', rows '1:3'). Failures are "E_CODE|detail" codes that excel.mjs turns into Korean;
# Invoke-ExcelEdit wraps them as "opfail|<n>|<op>|<message>" so the caller knows which operation failed.
# A COM Range is enumerable: functions return it as ",$rg" so PowerShell does not unroll it into cells.

function XlFail([string]$code, $detail) { throw ('{0}|{1}' -f $code, $detail) }

function XlProp($o, [string]$name) {
  if ($null -eq $o) { return $null }
  $p = $o.PSObject.Properties[$name]
  if ($p) { return $p.Value }
  return $null
}
function XlHas($o, [string]$name) { return ($null -ne $o -and $null -ne $o.PSObject.Properties[$name]) }

# 'RRGGBB' -> Excel BGR integer
function XlBgr([string]$hex) {
  $h = $hex.TrimStart('#')
  $r = [Convert]::ToInt32($h.Substring(0, 2), 16)
  $g = [Convert]::ToInt32($h.Substring(2, 2), 16)
  $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
  return [int]($r + 256 * $g + 65536 * $b)
}
function XlHex($bgr) {
  if ($null -eq $bgr) { return $null }
  $v = [int][double]$bgr
  return ('{0:X2}{1:X2}{2:X2}' -f ($v -band 255), (($v -shr 8) -band 255), (($v -shr 16) -band 255))
}

# sheet: name (case-insensitive like Excel) or 1-based index; nothing = first sheet
function XlSheet($wb, $ref) {
  if ($null -eq $ref -or "$ref" -eq '') { return $wb.Worksheets.Item(1) }
  if ($ref -is [string]) {
    foreach ($ws in $wb.Worksheets) { if ($ws.Name -eq $ref) { return $ws } }
    if ($ref -notmatch '^\d+$') { XlFail 'E_SHEET' $ref }
  }
  $i = [int]$ref
  if ($i -lt 1 -or $i -gt $wb.Worksheets.Count) { XlFail 'E_SHEET' $ref }
  return $wb.Worksheets.Item($i)
}

function XlRange($ws, [string]$addr) {
  $rg = $null
  try { $rg = $ws.Range($addr) } catch { XlFail 'E_RANGE' $addr }
  return , $rg
}

function XlAddr($rg) { return [string]$rg.Address($false, $false) }

function XlSheetNameTaken($wb, [string]$name) {
  foreach ($s in $wb.Sheets) { if ($s.Name -eq $name) { return $true } }
  return $false
}

function XlVisibleCount($wb) {
  $n = 0
  foreach ($s in $wb.Sheets) { if ($s.Visible -eq -1) { $n++ } }
  return $n
}

# writing into a merged area is only allowed at its top-left cell (same rule as the file engine)
function XlCheckNotMergedPart($cell) {
  if ($cell.MergeCells) {
    $tl = $cell.MergeArea.Cells.Item(1, 1)
    if ($tl.Row -ne $cell.Row -or $tl.Column -ne $cell.Column) { XlFail 'E_MERGED_PART' ((XlAddr $cell) + '|' + (XlAddr $tl)) }
  }
}

# inserting rows/columns would push used cells off the end of the sheet
function XlCheckRoom($ws, [string]$axis, [int]$at, [int]$n) {
  $ur = $ws.UsedRange
  if ([int]$ur.Rows.Count -eq 1 -and [int]$ur.Columns.Count -eq 1 -and $null -eq $ur.Value2) { return }
  if ($axis -eq 'row') {
    $lastUsed = [int]$ur.Row + [int]$ur.Rows.Count - 1
    if ($lastUsed -ge $at -and $lastUsed + $n -gt 1048576) { XlFail 'E_PUSH_OFF' 'row' }
  } else {
    $lastUsed = [int]$ur.Column + [int]$ur.Columns.Count - 1
    if ($lastUsed -ge $at -and $lastUsed + $n -gt 16384) { XlFail 'E_PUSH_OFF' 'col' }
  }
}

# a table or defined name already uses this name (case-insensitive like Excel)
function XlNameTaken($wb, [string]$name) {
  foreach ($s in $wb.Worksheets) {
    foreach ($lo in $s.ListObjects) { if ([string]$lo.Name -eq $name -or [string]$lo.DisplayName -eq $name) { return $true } }
  }
  foreach ($n in $wb.Names) {
    $nm = [string]$n.Name
    $nm = $nm.Substring($nm.LastIndexOf('!') + 1)
    if ($nm -eq $name) { return $true }
  }
  return $false
}

# PowerShell's COM adapter remembers the type of the first value set on a property (Value2 = 'text' then
# Value2 = object[,] fails): set variant properties through reflection instead
function XlSet($obj, [string]$prop, $value) {
  if ($value -is [System.Management.Automation.PSObject]) { $value = $value.PSObject.BaseObject }
  [void][System.__ComObject].InvokeMember($prop, [System.Reflection.BindingFlags]::SetProperty, $null, $obj, [object[]]@(, $value))
}

function XlGet($obj, [string]$prop) {
  return [System.__ComObject].InvokeMember($prop, [System.Reflection.BindingFlags]::GetProperty, $null, $obj, $null)
}

function XlSetFormula($rg, [string]$f) {
  try { XlSet $rg 'Formula2' $f; return } catch {}
  try { XlSet $rg 'Formula' $f } catch { XlFail 'E_FORMULA' $f }
}
function XlSetFormulas($rg, [object[,]]$arr) {
  try { XlSet $rg 'Formula2' $arr; return } catch {}
  try { XlSet $rg 'Formula' $arr } catch { XlFail 'E_FORMULA' (XlAddr $rg) }
}

# JSON value -> cell. Strings stay text (Excel would turn '120' or '1/2' into numbers: re-write those with a quote prefix
# and put back the number format Excel picked for them, e.g. a date format for '1/2')
function XlSetCell($cell, $v) {
  if ($null -eq $v) {
    if ($cell.MergeCells) { [void]$cell.MergeArea.ClearContents() } else { [void]$cell.ClearContents() }
    return
  }
  if ($v -is [string]) {
    $s = [string]$v
    if ($s.StartsWith('=')) { XlSetFormula $cell $s; return }
    $nf = $null
    try { $nf = $cell.NumberFormat } catch {}
    XlSet $cell 'Value2' $s
    if (-not $s.StartsWith("'")) {
      $back = $cell.Value2
      if ($back -isnot [string]) {
        XlSet $cell 'Value2' ("'" + $s)
        if ($nf -is [string]) { try { $cell.NumberFormat = $nf } catch {} }
      }
    }
    return
  }
  if ($v -is [bool]) { XlSet $cell 'Value2' ([bool]$v); return }
  XlSet $cell 'Value2' ([double]$v)
}

# unwrap JSON values (PSObject-wrapped values do not convert to COM variants)
function XlJsonValue($v) {
  if ($null -eq $v) { return $null }
  if ($v -is [string]) { return [string]$v }
  if ($v -is [bool]) { return [bool]$v }
  return [double]$v
}

# ---------------------------------------------------------------- inspect

# merged areas: Find with a merge format (FindNext ignores SearchFormat, so Find again after the last hit)
function XlMergeCount($ws, $ur) {
  $m = $ur.MergeCells
  if ($m -eq $false) { return 0 }
  if ($m -eq $true -and $ur.MergeArea.Address() -eq $ur.Address()) { return 1 }
  $app = $ws.Application
  $seen = @{}
  try {
    [void]$app.FindFormat.Clear()
    $app.FindFormat.MergeCells = $true
    $first = $null; $after = $ur.Cells.Item($ur.Rows.Count, $ur.Columns.Count)
    for ($i = 0; $i -lt 100000; $i++) {
      # What, After, LookIn(xlFormulas), LookAt(xlPart), SearchOrder(xlByRows), Direction(xlNext), MatchCase, MatchByte, SearchFormat
      $c = $ur.Find('', $after, -4123, 2, 1, 1, $false, $false, $true)
      if ($null -eq $c) { break }
      $addr = $c.Address()
      if ($null -eq $first) { $first = $addr } elseif ($addr -eq $first) { break }
      if ($c.MergeCells) { $seen[$c.MergeArea.Address()] = $true }
      $after = $c
    }
  } finally { try { [void]$app.FindFormat.Clear() } catch {} }
  return $seen.Count
}

function XlInspectCore($wb, $app) {
  # big workbooks: full details for the first $maxSheets sheets, the first $maxObjs tables/charts of each (counts stay exact)
  $maxFormulas = 20; $maxSheets = 100; $maxObjs = 20
  $active = $null
  try { $active = $wb.ActiveSheet } catch {}
  $win = $null
  try { $win = $wb.Windows.Item(1) } catch {}
  $sheets = New-Object System.Collections.ArrayList
  $formulas = New-Object System.Collections.ArrayList
  $formulaTotal = 0
  $index = 0
  foreach ($ws in $wb.Worksheets) {
    $index++
    $vis = 'visible'
    if ($ws.Visible -eq 0) { $vis = 'hidden' } elseif ($ws.Visible -eq 2) { $vis = 'veryHidden' }
    if ($index -gt $maxSheets) {
      [void]$sheets.Add(@{ index = $index; name = $ws.Name; visible = $vis; brief = $true })
      continue
    }
    $ur = $ws.UsedRange
    $empty = ($ur.Cells.Count -eq 1 -and $null -eq $ur.Value2 -and -not $ur.HasFormula)
    $info = @{ index = $index; name = $ws.Name; visible = $vis; usedRange = $null; rows = 0; cols = 0;
      freeze = $null; tables = @(); charts = @(); merges = 0; autoFilter = $null; conditionalFormats = 0; validations = 0 }
    if (-not $empty) {
      $info.usedRange = XlAddr $ur
      $info.rows = [int]$ur.Rows.Count; $info.cols = [int]$ur.Columns.Count
      try { $info.merges = [int](XlMergeCount $ws $ur) } catch { $info.merges = -1 }
    }
    # freeze panes live on the window: look at each visible sheet once (the workbook is opened read-only)
    if ($win -and $vis -eq 'visible') {
      try {
        [void]$ws.Activate()
        if ($win.FreezePanes) {
          $p = $win.Panes.Item(1)
          $info.freeze = (XlAddr $ws.Cells.Item($p.ScrollRow + $win.SplitRow, $p.ScrollColumn + $win.SplitColumn))
        }
      } catch {}
    }
    $tables = New-Object System.Collections.ArrayList
    $info.tableCount = [int]$ws.ListObjects.Count
    foreach ($lo in $ws.ListObjects) {
      if ($tables.Count -ge $maxObjs) { break }
      $style = $null
      try { $style = $lo.TableStyle.Name } catch { try { $style = [string]$lo.TableStyle } catch {} }
      [void]$tables.Add(@{ name = $lo.Name; range = (XlAddr $lo.Range); style = $style; totals = [bool]$lo.ShowTotals })
    }
    $info.tables = @($tables)
    $charts = New-Object System.Collections.ArrayList
    $cos = $ws.ChartObjects()
    $info.chartCount = [int]$cos.Count
    foreach ($co in $cos) {
      if ($charts.Count -ge $maxObjs) { break }
      $title = $null
      try { if ($co.Chart.HasTitle) { $title = $co.Chart.ChartTitle.Text } } catch {}
      [void]$charts.Add(@{ name = $co.Name; type = [int]$co.Chart.ChartType; title = $title })
    }
    $info.charts = @($charts)
    if ($ws.AutoFilterMode) { try { $info.autoFilter = XlAddr $ws.AutoFilter.Range } catch {} }
    try { $info.conditionalFormats = [int]$ws.Cells.FormatConditions.Count } catch {}
    try { $info.validations = [int]$ws.Cells.SpecialCells(-4174).Areas.Count } catch {}
    if (-not $empty) {
      $fc = $null
      try { $fc = $ur.SpecialCells(-4123) } catch {}
      if ($fc) {
        $formulaTotal += [int]$fc.CountLarge
        :areas foreach ($area in $fc.Areas) {
          foreach ($c in $area.Cells) {
            if ($formulas.Count -ge $maxFormulas) { break areas }
            $f = $null
            try { $f = XlGet $c 'Formula2' } catch { $f = XlGet $c 'Formula' }
            [void]$formulas.Add(@{ sheet = $ws.Name; cell = (XlAddr $c); formula = [string]$f })
          }
        }
      }
    }
    [void]$sheets.Add($info)
  }
  if ($active) { try { [void]$active.Activate() } catch {} }
  $names = New-Object System.Collections.ArrayList
  $hiddenNames = 0
  foreach ($n in $wb.Names) {
    if (-not $n.Visible) { $hiddenNames++; continue }
    $scope = $null; $name = [string]$n.Name
    try { if ($n.Parent.Name -ne $wb.Name) { $scope = [string]$n.Parent.Name; $name = $name.Substring($name.LastIndexOf('!') + 1) } } catch {}
    [void]$names.Add(@{ name = $name; refersTo = [string]$n.RefersTo; scope = $scope })
  }
  return @{
    activeSheet = $(if ($active) { [string]$active.Name } else { $null })
    sheets = @($sheets); names = @($names); hiddenNames = $hiddenNames
    formulas = @($formulas); formulaCount = $formulaTotal
  }
}

# ---------------------------------------------------------------- edit

$XlChartTypes = @{ column = 51; bar = 57; line = 4; pie = 5; scatter = -4169; area = 1 }
$XlHAlign = @{ general = 1; left = -4131; center = -4108; right = -4152; justify = -4130; center_across = 7; fill = 5 }
$XlVAlign = @{ top = -4160; center = -4108; bottom = -4107; justify = -4130 }
$XlWeight = @{ thin = 2; medium = -4138; thick = 4 }

function XlBorder($rg, [string]$kind) {
  switch ($kind) {
    'none' { foreach ($i in 5..12) { try { $rg.Borders.Item($i).LineStyle = -4142 } catch {} } }
    'outline' { [void]$rg.BorderAround(1, 2) }
    default {
      foreach ($i in 7..12) {
        try { $b = $rg.Borders.Item($i); $b.LineStyle = 1; $b.Weight = $XlWeight[$kind] } catch {}
      }
    }
  }
}

# literal list for Validation.Add uses the list separator of this Excel
function XlListFormula($app, $items) {
  $sep = ','
  try { $s = [string]$app.International(5); if ($s) { $sep = $s } } catch {}
  return (@($items) | ForEach-Object { [string]$_ }) -join $sep
}

# number -> '=10', text -> '="abc"', '=...' stays a formula
function XlCfValue($v) {
  if ($v -is [string]) {
    if ($v.StartsWith('=')) { return $v }
    return '="' + $v.Replace('"', '""') + '"'
  }
  return '=' + ([double]$v).ToString([System.Globalization.CultureInfo]::InvariantCulture)
}

function XlTableOverlap($ws, $rg) {
  foreach ($lo in $ws.ListObjects) {
    if ($null -ne $ws.Application.Intersect($lo.Range, $rg)) { return [string]$lo.Name }
  }
  return $null
}

function Invoke-XlOp($wb, $app, $op) {
  $kind = [string]$op.op
  $out = @{}
  switch ($kind) {
    'add_sheet' {
      $name = [string]$op.name
      if (XlSheetNameTaken $wb $name) { XlFail 'E_SHEET_EXISTS' $name }
      $after = $(if (XlHas $op 'after') { XlSheet $wb $op.after } else { $wb.Worksheets.Item($wb.Worksheets.Count) })
      $new = $wb.Worksheets.Add($Missing, $after)
      try { $new.Name = $name } catch { XlFail 'E_SHEET_NAME' $name }
      $out.sheet = $new.Name; $out.index = [int]$new.Index
    }
    'rename_sheet' {
      $ws = XlSheet $wb $op.sheet
      $name = [string]$op.name
      $old = $ws.Name
      if ($old -ne $name -and (XlSheetNameTaken $wb $name)) { XlFail 'E_SHEET_EXISTS' $name }
      try { $ws.Name = $name } catch { XlFail 'E_SHEET_NAME' $name }
      $out.from = $old; $out.sheet = $ws.Name
    }
    'delete_sheet' {
      $ws = XlSheet $wb $op.sheet
      if ($ws.Visible -eq -1 -and (XlVisibleCount $wb) -le 1) { XlFail 'E_LAST_SHEET' $ws.Name }
      $out.sheet = $ws.Name
      [void]$ws.Delete()
    }
    'copy_sheet' {
      $ws = XlSheet $wb $op.sheet
      $name = [string]$op.name
      if (XlSheetNameTaken $wb $name) { XlFail 'E_SHEET_EXISTS' $name }
      $idx = [int]$ws.Index
      [void]$ws.Copy($Missing, $ws)
      $new = $wb.Sheets.Item($idx + 1)
      try { $new.Name = $name } catch { XlFail 'E_SHEET_NAME' $name }
      $out.from = $ws.Name; $out.sheet = $new.Name; $out.index = [int]$new.Index
    }
    'move_sheet' {
      $ws = XlSheet $wb $op.sheet
      $count = $wb.Worksheets.Count
      $pos = [int]$op.position
      if ($pos -lt 1 -or $pos -gt $count) { XlFail 'E_POSITION' ('{0}|{1}' -f $pos, $count) }
      $cur = 0; $i = 0
      foreach ($s in $wb.Worksheets) { $i++; if ($s.Name -eq $ws.Name) { $cur = $i } }
      if ($pos -lt $cur) { [void]$ws.Move($wb.Worksheets.Item($pos)) }
      elseif ($pos -gt $cur) { [void]$ws.Move($Missing, $wb.Worksheets.Item($pos)) }
      $out.sheet = $ws.Name; $out.position = $pos
    }
    'set_cells' {
      $ws = XlSheet $wb $op.sheet
      $n = 0
      foreach ($p in $op.cells.PSObject.Properties) {
        $cell = XlRange $ws $p.Name
        XlCheckNotMergedPart $cell
        XlSetCell $cell (XlJsonValue $p.Value)
        $n++
      }
      $out.sheet = $ws.Name; $out.count = $n
    }
    'set_range' {
      $ws = XlSheet $wb $op.sheet
      $start = (XlRange $ws $op.range).Cells.Item(1, 1)
      $rows = @($op.values)
      $nr = $rows.Count; $nc = 1
      foreach ($row in $rows) { if (@($row).Count -gt $nc) { $nc = @($row).Count } }
      $target = $start.Resize($nr, $nc)
      $merged = ($target.MergeCells -ne $false)
      $arr = New-Object 'object[,]' $nr, $nc
      $hasFormula = $false; $hasText = $false
      for ($r = 0; $r -lt $nr; $r++) {
        $row = @($rows[$r])
        for ($c = 0; $c -lt $nc; $c++) {
          $v = $null
          if ($c -lt $row.Count) {
            # Excel drops values written into a merged area outside its top-left cell: refuse like set_cells
            if ($merged) { XlCheckNotMergedPart $target.Cells.Item($r + 1, $c + 1) }
            $v = XlJsonValue $row[$c]
          }
          if ($null -eq $v) { continue }
          if ($v -is [string]) {
            if ($v.StartsWith('=')) { $hasFormula = $true } elseif (-not $v.StartsWith("'")) { $hasText = $true }
            $arr[$r, $c] = [string]$v
          }
          elseif ($v -is [bool]) { $arr[$r, $c] = [bool]$v }
          else { $arr[$r, $c] = [double]$v }
        }
      }
      # number formats before writing (Excel changes them for strings it reads as dates or percents)
      $nf0 = $null; $nfCells = $null
      if ($hasText) {
        try { $nf0 = $target.NumberFormat } catch {}
        if ($nf0 -isnot [string]) {
          $nfCells = @{}
          for ($r = 0; $r -lt $nr; $r++) {
            for ($c = 0; $c -lt $nc; $c++) {
              $v = $arr[$r, $c]
              if ($v -is [string] -and -not $v.StartsWith('=') -and -not $v.StartsWith("'")) {
                try { $nfCells[('{0},{1}' -f $r, $c)] = $target.Cells.Item($r + 1, $c + 1).NumberFormat } catch {}
              }
            }
          }
        }
      }
      if ($hasFormula) { XlSetFormulas $target $arr } else { XlSet $target 'Value2' $arr }
      # strings Excel turned into numbers/dates go back in as text, with their old number format
      if ($hasText) {
        $back = $target.Value2
        for ($r = 0; $r -lt $nr; $r++) {
          for ($c = 0; $c -lt $nc; $c++) {
            $v = $arr[$r, $c]
            if ($v -is [string] -and -not $v.StartsWith('=') -and -not $v.StartsWith("'")) {
              $b = $(if ($nr -eq 1 -and $nc -eq 1) { $back } else { $back[($r + 1), ($c + 1)] })
              if ($b -isnot [string]) {
                $cell = $target.Cells.Item($r + 1, $c + 1)
                XlSet $cell 'Value2' ([string]("'" + $v))
                $nf = $(if ($nfCells) { $nfCells[('{0},{1}' -f $r, $c)] } else { $nf0 })
                if ($nf -is [string]) { try { $cell.NumberFormat = $nf } catch {} }
              }
            }
          }
        }
      }
      $out.sheet = $ws.Name; $out.range = XlAddr $target; $out.rows = $nr; $out.cols = $nc
    }
    'clear' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.range
      try {
        switch ([string]$op.what) {
          'all' { [void]$rg.Clear() }
          'formats' { [void]$rg.ClearFormats() }
          default { [void]$rg.ClearContents() }
        }
      } catch {
        # only part of a merged area is in the range
        if ($rg.MergeCells -ne $false) { XlFail 'E_MERGE_PARTIAL' (XlAddr $rg) }
        throw
      }
      $out.sheet = $ws.Name; $out.range = XlAddr $rg
    }
    'insert_rows' {
      $ws = XlSheet $wb $op.sheet
      $at = [int]$op.at; $n = [int]$op.count
      XlCheckRoom $ws 'row' $at $n
      [void]$ws.Range(('{0}:{1}' -f $at, ($at + $n - 1))).EntireRow.Insert(-4121, 0)
      $out.sheet = $ws.Name
    }
    'delete_rows' {
      $ws = XlSheet $wb $op.sheet
      $at = [int]$op.at; $n = [int]$op.count
      $last = $at + $n - 1
      # a table's header or totals row can only go with the whole table (same rule as the file engine)
      foreach ($lo in $ws.ListObjects) {
        $t1 = [int]$lo.Range.Row; $t2 = $t1 + [int]$lo.Range.Rows.Count - 1
        if ($at -le $t1 -and $last -ge $t2) { continue }
        if ($lo.ShowHeaders -and $at -le $t1 -and $last -ge $t1) { XlFail 'E_TABLE_HEADER' ('{0}|{1}' -f $lo.Name, $t1) }
        if ($lo.ShowTotals -and $at -le $t2 -and $last -ge $t2) { XlFail 'E_TABLE_TOTALS' ('{0}|{1}' -f $lo.Name, $t2) }
      }
      [void]$ws.Range(('{0}:{1}' -f $at, $last)).EntireRow.Delete(-4162)
      $out.sheet = $ws.Name
    }
    'insert_columns' {
      $ws = XlSheet $wb $op.sheet
      XlCheckRoom $ws 'col' ([int]$op.at) ([int]$op.count)
      [void]$ws.Range([string]$op.columns).EntireColumn.Insert(-4161, 0)
      $out.sheet = $ws.Name
    }
    'delete_columns' {
      $ws = XlSheet $wb $op.sheet
      [void]$ws.Range([string]$op.columns).EntireColumn.Delete(-4159)
      $out.sheet = $ws.Name
    }
    'format' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.range
      $font = $rg.Font
      if (XlHas $op 'bold') { $font.Bold = [bool]$op.bold }
      if (XlHas $op 'italic') { $font.Italic = [bool]$op.italic }
      if (XlHas $op 'underline') { $font.Underline = $(if ($op.underline) { 2 } else { -4142 }) }
      if (XlHas $op 'font') { $font.Name = [string]$op.font }
      if (XlHas $op 'size') { $font.Size = [double]$op.size }
      if (XlHas $op 'color') { $font.Color = XlBgr $op.color }
      if (XlHas $op 'fill') {
        if ($op.fill -eq 'none') { $rg.Interior.Pattern = -4142 } else { $rg.Interior.Color = XlBgr $op.fill }
      }
      if (XlHas $op 'number_format') { $rg.NumberFormat = [string]$op.number_format }
      if (XlHas $op 'align') { $rg.HorizontalAlignment = $XlHAlign[[string]$op.align] }
      if (XlHas $op 'valign') { $rg.VerticalAlignment = $XlVAlign[[string]$op.valign] }
      if (XlHas $op 'wrap') { $rg.WrapText = [bool]$op.wrap }
      if (XlHas $op 'border') { XlBorder $rg ([string]$op.border) }
      $out.sheet = $ws.Name; $out.range = XlAddr $rg
    }
    'column_width' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.columns
      if ($op.width -eq 'auto') { [void]$rg.EntireColumn.AutoFit() } else { $rg.ColumnWidth = [double]$op.width }
      $out.sheet = $ws.Name
    }
    'row_height' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.rows
      if ($op.height -eq 'auto') { [void]$rg.EntireRow.AutoFit() } else { $rg.RowHeight = [double]$op.height }
      $out.sheet = $ws.Name
    }
    'merge' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.range
      $t = XlTableOverlap $ws $rg
      if ($t) { XlFail 'E_TABLE_OVERLAP' $t }
      # an existing merge sticking out of the range would silently grow the result
      if ($rg.MergeCells -ne $false) {
        $checked = @{}
        foreach ($c in $rg.Cells) {
          if (-not $c.MergeCells) { continue }
          $ma = $c.MergeArea
          $key = $ma.Address()
          if ($checked[$key]) { continue }
          $checked[$key] = $true
          if ($app.Intersect($ma, $rg).Cells.Count -ne $ma.Cells.Count) { XlFail 'E_MERGE_OVERLAP' (XlAddr $ma) }
        }
      }
      $filled = [int]$app.WorksheetFunction.CountA($rg)
      $tl = $rg.Cells.Item(1, 1)
      if ($null -ne $tl.Value2 -and "$($tl.Value2)" -ne '') { $filled-- }
      [void]$rg.Merge()
      $out.sheet = $ws.Name; $out.range = XlAddr $rg; $out.discarded = [Math]::Max(0, $filled)
    }
    'unmerge' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.range
      [void]$rg.UnMerge()
      $out.sheet = $ws.Name; $out.range = XlAddr $rg
    }
    'freeze' {
      $ws = XlSheet $wb $op.sheet
      $hidden = $ws.Visible -ne -1
      if ($hidden) { $ws.Visible = -1 }
      [void]$ws.Activate()
      $win = $wb.Windows.Item(1)
      $win.FreezePanes = $false
      $win.SplitRow = 0; $win.SplitColumn = 0
      if (XlProp $op 'cell') {
        $cell = XlRange $ws $op.cell
        $win.ScrollRow = 1; $win.ScrollColumn = 1
        if ($cell.Row -gt 1 -or $cell.Column -gt 1) {
          $win.SplitColumn = $cell.Column - 1
          $win.SplitRow = $cell.Row - 1
          $win.FreezePanes = $true
        }
        $out.cell = XlAddr $cell
      }
      if ($hidden) { $ws.Visible = 0 }
      $out.sheet = $ws.Name
    }
    'sort' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.range
      $col = $ws.Range(([string]$op.by) + '1').Column
      if ($col -lt $rg.Column -or $col -ge $rg.Column + $rg.Columns.Count) { XlFail 'E_SORT_KEY' ([string]$op.by) }
      $key = $ws.Cells.Item($rg.Row, $col)
      $order = $(if ($op.order -eq 'desc') { 2 } else { 1 })
      $header = 0
      if (XlHas $op 'header') { $header = $(if ($op.header) { 1 } else { 2 }) }
      # Key1, Order1, Key2, Type, Order2, Key3, Order3, Header
      try { [void]$rg.Sort($key, $order, $Missing, $Missing, 1, $Missing, 1, $header) }
      catch {
        if ($rg.MergeCells -ne $false) { XlFail 'E_SORT_MERGED' (XlAddr $rg) }
        throw
      }
      $out.sheet = $ws.Name; $out.range = XlAddr $rg
    }
    'autofilter' {
      $ws = XlSheet $wb $op.sheet
      if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }
      if (XlProp $op 'range') {
        $rg = XlRange $ws $op.range
        $t = XlTableOverlap $ws $rg
        if ($t) { XlFail 'E_TABLE_OVERLAP' $t }
        [void]$rg.AutoFilter()
        $out.range = XlAddr $ws.AutoFilter.Range
      }
      $out.sheet = $ws.Name
    }
    'add_table' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.range
      $t = XlTableOverlap $ws $rg
      if ($t) { XlFail 'E_TABLE_OVERLAP' $t }
      if ($rg.MergeCells -ne $false) { XlFail 'E_TABLE_MERGED' (XlAddr $rg) }
      # Excel quietly renames a table whose name is taken ('T' -> 'T_1'): refuse instead
      $want = [string](XlProp $op 'name')
      if ($want -and (XlNameTaken $wb $want)) { XlFail 'E_TABLE_EXISTS' $want }
      if ($ws.AutoFilterMode -and $null -ne $app.Intersect($ws.AutoFilter.Range, $rg)) { $ws.AutoFilterMode = $false }
      # SourceType xlSrcRange, Source, LinkSource, XlListObjectHasHeaders xlYes
      $lo = $ws.ListObjects.Add(1, $rg, $Missing, 1)
      if ($want) {
        try { $lo.Name = $want } catch { XlFail 'E_TABLE_NAME' $want }
        if ([string]$lo.DisplayName -ne $want) { XlFail 'E_TABLE_EXISTS' $want }
      }
      $style = $(if (XlProp $op 'style') { [string]$op.style } else { 'TableStyleMedium2' })
      try { $lo.TableStyle = $style } catch { XlFail 'E_TABLE_STYLE' $style }
      $out.sheet = $ws.Name; $out.name = [string]$lo.DisplayName; $out.range = XlAddr $lo.Range
    }
    'add_chart' {
      $ws = XlSheet $wb $op.sheet
      $src = XlRange $ws $op.range
      $at = $(if (XlProp $op 'at') { XlRange $ws $op.at } else { $ws.Cells.Item($src.Row, $src.Column + $src.Columns.Count + 1) })
      $w = $(if (XlProp $op 'width') { [double]$op.width } else { 360.0 })
      $h = $(if (XlProp $op 'height') { [double]$op.height } else { 216.0 })
      $type = $XlChartTypes[[string]$op.type]
      $shape = $null
      try { $shape = $ws.Shapes.AddChart2(-1, $type, [double]$at.Left, [double]$at.Top, $w, $h) }
      catch { $shape = $ws.Shapes.AddChart($type, [double]$at.Left, [double]$at.Top, $w, $h) }
      $chart = $shape.Chart
      # series follow Excel's Insert Chart rule (more rows than columns: one series per column) unless series_in says
      $plotBy = $(if ($op.series_in -eq 'rows') { 1 } elseif ($op.series_in -eq 'columns') { 2 }
        elseif ([int]$src.Rows.Count -gt [int]$src.Columns.Count) { 2 } else { 1 })
      [void]$chart.SetSourceData($src, $plotBy)
      $nSeries = [int]$chart.SeriesCollection().Count
      if (XlProp $op 'title') { $chart.HasTitle = $true; $chart.ChartTitle.Text = [string]$op.title }
      else {
        # no title asked: keep only Excel's automatic series-name title, never the 'Chart Title' placeholder
        $keep = $false
        try { if ($nSeries -eq 1 -and $chart.HasTitle) { $keep = ([string]$chart.ChartTitle.Text -eq [string]$chart.SeriesCollection(1).Name) } } catch {}
        if (-not $keep) { try { $chart.HasTitle = $false } catch {} }
      }
      $out.sheet = $ws.Name; $out.name = $shape.Name; $out.range = XlAddr $src; $out.at = XlAddr $at
      $out.series = $nSeries; $out.seriesIn = $(if ($plotBy -eq 2) { 'columns' } else { 'rows' })
      $out.title = $null
      try { if ($chart.HasTitle) { $out.title = [string]$chart.ChartTitle.Text } } catch {}
    }
    'conditional_format' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.range
      $rule = $op.rule
      $type = [string]$rule.type
      $fc = $null
      switch ($type) {
        'greater_than' { $fc = $rg.FormatConditions.Add(1, 5, (XlCfValue $rule.value)) }
        'less_than' { $fc = $rg.FormatConditions.Add(1, 6, (XlCfValue $rule.value)) }
        'equal' { $fc = $rg.FormatConditions.Add(1, 3, (XlCfValue $rule.value)) }
        'between' { $fc = $rg.FormatConditions.Add(1, 1, (XlCfValue $rule.value), (XlCfValue $rule.value2)) }
        'text_contains' { $fc = $rg.FormatConditions.Add(9, $Missing, $Missing, $Missing, [string]$rule.value, 0) }
        'duplicate' { $fc = $rg.FormatConditions.AddUniqueValues(); $fc.DupeUnique = 1 }
        'color_scale' {
          $fc = $rg.FormatConditions.AddColorScale(3)
          $colors = @('F8696B', 'FFEB84', '63BE7B')
          for ($i = 1; $i -le 3; $i++) { $fc.ColorScaleCriteria.Item($i).FormatColor.Color = XlBgr $colors[$i - 1] }
        }
        'data_bar' {
          $fc = $rg.FormatConditions.AddDatabar()
          $fc.BarColor.Color = XlBgr $(if (XlProp $rule 'fill') { [string]$rule.fill } else { '638EC6' })
        }
      }
      if ($type -notin @('color_scale', 'data_bar')) {
        $fill = $(if (XlHas $rule 'fill') { XlProp $rule 'fill' } elseif (XlHas $rule 'font_color') { $null } else { 'FFC7CE' })
        $fontColor = $(if (XlHas $rule 'font_color') { XlProp $rule 'font_color' } elseif (XlHas $rule 'fill') { $null } else { '9C0006' })
        if ($fill) { $fc.Interior.Color = XlBgr $fill }
        if ($fontColor) { $fc.Font.Color = XlBgr $fontColor }
      }
      $out.sheet = $ws.Name; $out.range = XlAddr $rg
    }
    'data_validation' {
      $ws = XlSheet $wb $op.sheet
      $rg = XlRange $ws $op.range
      $v = $rg.Validation
      [void]$v.Delete()
      if (XlHas $op 'list') {
        $list = $op.list
        $f = $(if ($list -is [string]) { [string]$list } else { XlListFormula $app $list })
        # Type xlValidateList, AlertStyle xlValidAlertStop, Operator xlBetween, Formula1
        [void]$v.Add(3, 1, 1, $f)
        $v.InCellDropdown = $true
      } else {
        $vt = $(if ($op.type -eq 'decimal') { 2 } else { 1 })
        $hasMin = XlHas $op 'min'; $hasMax = XlHas $op 'max'
        $inv = [System.Globalization.CultureInfo]::InvariantCulture
        if ($hasMin -and $hasMax) { [void]$v.Add($vt, 1, 1, ([double]$op.min).ToString($inv), ([double]$op.max).ToString($inv)) }
        elseif ($hasMin) { [void]$v.Add($vt, 1, 7, ([double]$op.min).ToString($inv)) }
        else { [void]$v.Add($vt, 1, 8, ([double]$op.max).ToString($inv)) }
      }
      $v.IgnoreBlank = $true
      $v.ShowError = $true
      $out.sheet = $ws.Name; $out.range = XlAddr $rg
    }
    'named_range' {
      $ws = XlSheet $wb $op.sheet
      $refersTo = "='" + $ws.Name.Replace("'", "''") + "'!" + [string]$op.area
      $n = $null
      try { $n = $wb.Names.Add([string]$op.name, $refersTo) } catch { XlFail 'E_NAME' ([string]$op.name) }
      $out.name = [string]$op.name; $out.refersTo = [string]$n.RefersTo
    }
    'recalculate' { [void]$app.CalculateFull() }
    default { XlFail 'E_UNKNOWN_OP' $kind }
  }
  return $out
}

function XlEditCore($wb, $app, $req) {
  # Excel opens a file another program holds as read-only (alerts off) and Save() then writes nothing
  if ($req.save -and ([string]$req.save -eq [string]$req.path) -and $wb.ReadOnly) { throw ('locked for editing: ' + [string]$req.path) }
  $ops = @($req.operations)
  $results = New-Object System.Collections.ArrayList
  $active = $null
  try { $active = $wb.ActiveSheet } catch {}
  for ($i = 0; $i -lt $ops.Count; $i++) {
    $op = $ops[$i]
    try { $r = Invoke-XlOp $wb $app $op }
    catch { throw ('opfail|{0}|{1}|{2}' -f ($i + 1), [string]$op.op, $_.Exception.Message) }
    [void]$results.Add($r)
  }
  # csv keeps only the active sheet: show the one the operations wrote to (all must be the same sheet)
  $csvSheet = $null
  if (XlHas $req 'csvSheets') {
    foreach ($ref in @($req.csvSheets)) {
      $s = XlSheet $wb $ref
      if ($csvSheet -and $csvSheet.Name -ne $s.Name) { XlFail 'E_CSV_SHEETS' ($csvSheet.Name + '|' + $s.Name) }
      $csvSheet = $s
    }
  }
  if ($csvSheet) {
    if ($csvSheet.Visible -ne -1) { $csvSheet.Visible = -1 }
    [void]$csvSheet.Activate()
  } elseif ($active) {
    # keep the sheet that was showing (adding or freezing sheets activates them)
    try { [void]$active.Activate() } catch {}
  }
  # make Save() write even when nothing Excel tracks changed
  if ($req.save) { try { $wb.Saved = $false } catch {} }
  return $results.ToArray()
}

# Entry points for office.ps1. The work runs in an inner function so its COM wrappers (ranges, windows,
# AutoFilter objects ...) are out of scope here and can be collected before office.ps1 closes and quits:
# a wrapper still alive at Quit can keep a hidden EXCEL.EXE running after the request.
function XlCollect { [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }

function Invoke-ExcelInspect($wb, $app, $req) {
  try { return (XlInspectCore $wb $app) } finally { XlCollect }
}

function Invoke-ExcelEdit($wb, $app, $req) {
  try { return (XlEditCore $wb $app $req) } finally { XlCollect }
}
