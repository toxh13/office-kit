# make-icon.ps1 - draws the Office Kit app icon (build/icon.png, build/icon.ico).
# Rounded square + document + sparkle. Each size is drawn separately so small icons stay crisp.
# (ASCII-only comments: Windows PowerShell 5.1 reads BOM-less scripts in the ANSI code page.)
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/make-icon.ps1
param([string]$OutDir = (Join-Path $PSScriptRoot '..\..\build'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function New-RoundRect([single]$x, [single]$y, [single]$w, [single]$h, [single]$r) {
  $p = New-Object System.Drawing.Drawing2D.GraphicsPath
  $d = $r * 2
  $p.AddArc($x, $y, $d, $d, 180, 90)
  $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
  $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
  $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
  $p.CloseFigure()
  return $p
}

# Four-pointed sparkle with concave sides
function New-Sparkle([single]$cx, [single]$cy, [single]$r) {
  $k = $r * 0.16
  $p = New-Object System.Drawing.Drawing2D.GraphicsPath
  $pts = @(
    @($cx, ($cy - $r)), @(($cx + $r), $cy), @($cx, ($cy + $r)), @(($cx - $r), $cy)
  )
  $ctl = @(
    @(($cx + $k), ($cy - $k)), @(($cx + $k), ($cy + $k)), @(($cx - $k), ($cy + $k)), @(($cx - $k), ($cy - $k))
  )
  for ($i = 0; $i -lt 4; $i++) {
    $a = $pts[$i]; $b = $pts[($i + 1) % 4]; $c = $ctl[$i]
    $p.AddBezier([single]$a[0], [single]$a[1], [single]$c[0], [single]$c[1], [single]$c[0], [single]$c[1], [single]$b[0], [single]$b[1])
  }
  $p.CloseFigure()
  return $p
}

function Color([string]$hex, [int]$alpha = 255) {
  $c = [System.Drawing.ColorTranslator]::FromHtml($hex)
  return [System.Drawing.Color]::FromArgb($alpha, $c.R, $c.G, $c.B)
}

# Drawn in a 256-unit space, then scaled to $size
function Draw-Icon([int]$size) {
  $bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $g.Clear([System.Drawing.Color]::Transparent)
  $s = $size / 256.0
  $g.ScaleTransform($s, $s)
  $small = $size -le 24

  # Background: indigo -> sky blue
  $m = $(if ($small) { 4 } else { 8 })
  $bg = New-RoundRect $m $m (256 - 2 * $m) (256 - 2 * $m) 56
  $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.PointF(0, 0)), (New-Object System.Drawing.PointF(256, 256)), (Color '#4338CA'), (Color '#0EA5E9'))
  $g.FillPath($grad, $bg)
  # Soft highlight at the top
  $shine = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.PointF(0, -2)), (New-Object System.Drawing.PointF(0, 258)), (Color '#FFFFFF' 46), (Color '#FFFFFF' 0))
  $blend = New-Object System.Drawing.Drawing2D.Blend(3)
  $blend.Positions = [single[]]@(0, 0.55, 1)
  $blend.Factors = [single[]]@(0, 1, 1)
  $shine.Blend = $blend
  $g.FillPath($shine, $bg)

  # Document with a folded top-right corner
  $dx = 62; $dy = 40; $dw = 120; $dh = 176; $fold = 36; $r = 14
  if ($small) { $dx = 52; $dy = 30; $dw = 140; $dh = 196; $fold = 44 }
  $doc = New-Object System.Drawing.Drawing2D.GraphicsPath
  $doc.AddLine($dx + $r, $dy, $dx + $dw - $fold, $dy)
  $doc.AddLine($dx + $dw - $fold, $dy, $dx + $dw, $dy + $fold)
  $doc.AddLine($dx + $dw, $dy + $fold, $dx + $dw, $dy + $dh - $r)
  $doc.AddArc($dx + $dw - 2 * $r, $dy + $dh - 2 * $r, 2 * $r, 2 * $r, 0, 90)
  $doc.AddArc($dx, $dy + $dh - 2 * $r, 2 * $r, 2 * $r, 90, 90)
  $doc.AddArc($dx, $dy, 2 * $r, 2 * $r, 180, 90)
  $doc.CloseFigure()
  if (-not $small) {
    $shadow = $doc.Clone()
    $mx = New-Object System.Drawing.Drawing2D.Matrix
    $mx.Translate(0, 6)
    $shadow.Transform($mx)
    $g.FillPath((New-Object System.Drawing.SolidBrush((Color '#1E1B4B' 70))), $shadow)
  }
  $g.FillPath((New-Object System.Drawing.SolidBrush((Color '#FFFFFF'))), $doc)
  $foldPath = New-Object System.Drawing.Drawing2D.GraphicsPath
  $foldPath.AddLine($dx + $dw - $fold, $dy, $dx + $dw - $fold, $dy + $fold - 8)
  $foldPath.AddArc($dx + $dw - $fold, $dy + $fold - 16, 16, 16, 180, -90)
  $foldPath.AddLine($dx + $dw - $fold + 8, $dy + $fold, $dx + $dw, $dy + $fold)
  $foldPath.CloseFigure()
  $g.FillPath((New-Object System.Drawing.SolidBrush((Color '#C7D2FE'))), $foldPath)

  # Text lines
  $lineBrush = New-Object System.Drawing.SolidBrush((Color '#A5B4FC'))
  if ($small) {
    foreach ($ly in @(98, 140)) { $g.FillPath($lineBrush, (New-RoundRect ($dx + 22) $ly 84 22 11)) }
  } else {
    $lines = @(@(92, 64), @(118, 84), @(144, 84), @(170, 40))
    foreach ($l in $lines) { $g.FillPath($lineBrush, (New-RoundRect ($dx + 20) $l[0] $l[1] 12 6)) }
  }

  # Sparkle (AI), white outline separates it from the page
  $big = $(if ($small) { New-Sparkle 186 186 58 } else { New-Sparkle 184 180 48 })
  $pen = New-Object System.Drawing.Pen((Color '#FFFFFF'), $(if ($small) { 14 } else { 12 }))
  $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
  $g.DrawPath($pen, $big)
  $gold = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.PointF(120, 116)), (New-Object System.Drawing.PointF(250, 246)), (Color '#FDE68A'), (Color '#F59E0B'))
  $g.FillPath($gold, $big)
  if ($size -ge 48) {
    $tiny = New-Sparkle 214 100 15
    $g.DrawPath((New-Object System.Drawing.Pen((Color '#FFFFFF'), 6)), $tiny)
    $g.FillPath((New-Object System.Drawing.SolidBrush((Color '#FDE68A'))), $tiny)
  }
  $g.Dispose()
  return $bmp
}

function Get-PngBytes($bmp) {
  $ms = New-Object System.IO.MemoryStream
  $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
  return $ms.ToArray()
}

# Sizes < 256 are stored as BMP (DIB) entries, 256 as PNG (works for NSIS and Explorer)
function Get-DibBytes($bmp) {
  $w = $bmp.Width; $h = $bmp.Height
  $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
  $data = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $stride = $data.Stride
  $raw = New-Object byte[] ($stride * $h)
  [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $raw, 0, $raw.Length)
  $bmp.UnlockBits($data)
  $ms = New-Object System.IO.MemoryStream
  $bw = New-Object System.IO.BinaryWriter($ms)
  $bw.Write([int]40); $bw.Write([int]$w); $bw.Write([int]($h * 2)); $bw.Write([int16]1); $bw.Write([int16]32)
  $bw.Write([int]0); $bw.Write([int]($w * $h * 4)); $bw.Write([int]0); $bw.Write([int]0); $bw.Write([int]0); $bw.Write([int]0)
  for ($y = $h - 1; $y -ge 0; $y--) { $bw.Write($raw, $y * $stride, $w * 4) }   # bottom-up rows
  $maskRow = [int]([Math]::Ceiling($w / 32.0) * 4)
  $bw.Write((New-Object byte[] ($maskRow * $h)))                                   # AND mask (all zero, alpha is used)
  $bw.Flush()
  return $ms.ToArray()
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path

# Large PNG (README, other platforms)
$png = Draw-Icon 512
$png.Save((Join-Path $OutDir 'icon.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$png.Dispose()

# ICO
$sizes = @(16, 20, 24, 32, 40, 48, 64, 128, 256)
$images = @()
foreach ($sz in $sizes) {
  $b = Draw-Icon $sz
  $bytes = $(if ($sz -ge 256) { Get-PngBytes $b } else { Get-DibBytes $b })
  $images += , @($sz, $bytes)
  $b.Dispose()
}
$fs = [System.IO.File]::Create((Join-Path $OutDir 'icon.ico'))
$bw = New-Object System.IO.BinaryWriter($fs)
$bw.Write([int16]0); $bw.Write([int16]1); $bw.Write([int16]$images.Count)
$offset = 6 + 16 * $images.Count
foreach ($im in $images) {
  $sz = $im[0]; $len = $im[1].Length
  $dim = $(if ($sz -ge 256) { 0 } else { $sz })
  $bw.Write([byte]$dim); $bw.Write([byte]$dim); $bw.Write([byte]0); $bw.Write([byte]0)
  $bw.Write([int16]1); $bw.Write([int16]32); $bw.Write([int]$len); $bw.Write([int]$offset)
  $offset += $len
}
foreach ($im in $images) { $bw.Write([byte[]]$im[1]) }
$bw.Close()
Write-Output "icon.png, icon.ico -> $OutDir"
