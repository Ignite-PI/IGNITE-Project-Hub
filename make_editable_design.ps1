param(
    [Parameter(Mandatory = $true)][string]$Image,
    [string]$OutFile
)
# Turns a flat newsletter design (e.g. a Canva PNG/JPG) into an editable design for Project_Hub.html:
# Windows OCR finds the text, the text is painted out of the picture with the surrounding colour,
# and an editable text box (position, size, colour, bold) is created for every line / paragraph.
# Upload the resulting *_editable.json with "Upload new format" on the Newsletter page.
$ErrorActionPreference = 'Stop'
$Image = (Resolve-Path $Image).Path
if (-not $OutFile) { $OutFile = [IO.Path]::Combine([IO.Path]::GetDirectoryName($Image), [IO.Path]::GetFileNameWithoutExtension($Image) + '_editable.json') }
$DESIGN_W = 900       # design width used by Project_Hub.html
Add-Type -AssemblyName System.Drawing
# fast pixel work (finding pictures/icons and cutting them out) in C#
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
public static class DesignOps {
    static int D(byte[] p, int i, int r, int g, int b) { int dr = p[i + 2] - r, dg = p[i + 1] - g, db = p[i] - b; return (int)Math.Sqrt(dr * dr + dg * dg + db * db); }
    static int Cl(int v, int max) { return v < 0 ? 0 : (v > max ? max : v); }
    public static void Fill(byte[] p, int stride, int x0, int y0, int x1, int y1, int r, int g, int b) {
        for (int y = y0; y <= y1; y++) { int i = y * stride + x0 * 4; for (int x = x0; x <= x1; x++, i += 4) { p[i] = (byte)b; p[i + 1] = (byte)g; p[i + 2] = (byte)r; p[i + 3] = 255; } }
    }
    // true when nothing but plain background lies between two pieces of text (no border / picture in between)
    public static bool GapEmpty(byte[] p, int stride, int w, int h, int x0, int y0, int x1, int y1, int thr) {
        x0 = Cl(x0, w - 1); x1 = Cl(x1, w - 1); y0 = Cl(y0, h - 1); y1 = Cl(y1, h - 1);
        if (x1 <= x0) return true;
        int ri = y0 * stride + x0 * 4, r = p[ri + 2], g = p[ri + 1], b = p[ri];
        for (int y = y0; y <= y1; y += 2) for (int x = x0; x <= x1; x += 2) if (D(p, y * stride + x * 4, r, g, b) > thr) return false;
        return true;
    }
    // {bgR,bgG,bgB, fgR,fgG,fgB, inkRatio} for a line of text: background = median around the box, text = most different pixels inside
    public static double[] LineStats(byte[] p, int stride, int w, int h, int x0, int y0, int x1, int y1, int pad) {
        List<int>[] e = { new List<int>(), new List<int>(), new List<int>() };
        int ex0 = Cl(x0 - pad, w - 1), ey0 = Cl(y0 - pad, h - 1), ex1 = Cl(x1 + pad, w - 1), ey1 = Cl(y1 + pad, h - 1);
        for (int x = ex0; x <= ex1; x++) foreach (int y in new[] { ey0, ey1 }) { int i = y * stride + x * 4; e[0].Add(p[i + 2]); e[1].Add(p[i + 1]); e[2].Add(p[i]); }
        for (int y = ey0; y <= ey1; y++) foreach (int x in new[] { ex0, ex1 }) { int i = y * stride + x * 4; e[0].Add(p[i + 2]); e[1].Add(p[i + 1]); e[2].Add(p[i]); }
        int[] bg = new int[3]; for (int k = 0; k < 3; k++) { e[k].Sort(); bg[k] = e[k][e[k].Count / 2]; }
        // how "busy" the surrounding is (average spread of the colours around the line) - high = text sits on a photo
        double spread = 0; for (int k = 0; k < 3; k++) { double m = 0; foreach (int v in e[k]) m += v; m /= e[k].Count; double s2 = 0; foreach (int v in e[k]) s2 += (v - m) * (v - m); spread += Math.Sqrt(s2 / e[k].Count); } spread /= 3;
        int cx0 = Cl(x0, w - 1), cy0 = Cl(y0, h - 1), cx1 = Cl(x1, w - 1), cy1 = Cl(y1, h - 1), maxD = 0, n = 0;
        for (int y = cy0; y <= cy1; y++) for (int x = cx0; x <= cx1; x++) { int d = D(p, y * stride + x * 4, bg[0], bg[1], bg[2]); if (d > maxD) maxD = d; n++; }
        double sr = 0, sg = 0, sb = 0; int cnt = 0, heavy = 0;
        for (int y = cy0; y <= cy1; y++) for (int x = cx0; x <= cx1; x++) {
            int i = y * stride + x * 4, d = D(p, i, bg[0], bg[1], bg[2]);
            if (d >= maxD * 0.6) { sr += p[i + 2]; sg += p[i + 1]; sb += p[i]; cnt++; }
            if (d >= maxD * 0.45) heavy++;
        }
        double[] fg = (maxD < 40 || cnt == 0) ? new double[] { 0, 0, 0 } : new double[] { sr / cnt, sg / cnt, sb / cnt };
        return new double[] { bg[0], bg[1], bg[2], fg[0], fg[1], fg[2], n > 0 ? (double)heavy / n : 0, spread };
    }
    // share of pixels in the box that differ from the colour (bullet / leftover mark detection)
    public static double InkShare(byte[] p, int stride, int w, int h, int x0, int y0, int x1, int y1, int r, int g, int b, int thr) {
        x0 = Cl(x0, w - 1); x1 = Cl(x1, w - 1); y0 = Cl(y0, h - 1); y1 = Cl(y1, h - 1); if (x1 <= x0 || y1 <= y0) return 0;
        int n = 0, ink = 0;
        for (int y = y0; y <= y1; y++) for (int x = x0; x <= x1; x++) { n++; if (D(p, y * stride + x * 4, r, g, b) > thr) ink++; }
        return n > 0 ? (double)ink / n : 0;
    }
    // connected areas that differ from the page colour -> {id, minX, minY, maxX, maxY, pixelCount}
    public static int[] Label(byte[] p, int stride, int w, int h, int r, int g, int b, int thr, List<int[]> comps) {
        int[] lab = new int[w * h]; int n = 0; Stack<int> st = new Stack<int>();
        for (int y = 0; y < h; y++) for (int x = 0; x < w; x++) {
            int idx = y * w + x; if (lab[idx] != 0) continue;
            if (D(p, y * stride + x * 4, r, g, b) <= thr) { lab[idx] = -1; continue; }
            n++; int minx = x, maxx = x, miny = y, maxy = y, cnt = 0; lab[idx] = n; st.Push(idx);
            while (st.Count > 0) {
                int c = st.Pop(), cx = c % w, cy = c / w; cnt++;
                if (cx < minx) minx = cx; if (cx > maxx) maxx = cx; if (cy < miny) miny = cy; if (cy > maxy) maxy = cy;
                for (int dy = -1; dy <= 1; dy++) for (int dx = -1; dx <= 1; dx++) {
                    int nx = cx + dx, ny = cy + dy; if (nx < 0 || ny < 0 || nx >= w || ny >= h) continue;
                    int ni = ny * w + nx; if (lab[ni] != 0) continue;
                    if (D(p, ny * stride + nx * 4, r, g, b) <= thr) { lab[ni] = -1; continue; }
                    lab[ni] = n; st.Push(ni);
                }
            }
            comps.Add(new int[] { n, minx, miny, maxx, maxy, cnt });
        }
        return lab;
    }
    // copies the kept areas (plus their soft edge) inside the box to a transparent BGRA image and paints them with the page colour
    public static byte[] Extract(byte[] p, int stride, int[] lab, int w, bool[] keep, int x0, int y0, int x1, int y1, int r, int g, int b) {
        int cw = x1 - x0 + 1, ch = y1 - y0 + 1; byte[] o = new byte[cw * ch * 4];
        int h = lab.Length / w;
        for (int y = 0; y < ch; y++) for (int x = 0; x < cw; x++) {
            int gx = x + x0, gy = y + y0, li = gy * w + gx, l = lab[li];
            bool mine = l > 0 && keep[l];
            if (!mine && l == -1) { // soft edge pixel next to a kept area
                for (int dy = -1; dy <= 1 && !mine; dy++) for (int dx = -1; dx <= 1 && !mine; dx++) {
                    int nx = gx + dx, ny = gy + dy; if (nx < 0 || ny < 0 || nx >= w || ny >= h) continue;
                    int nl = lab[ny * w + nx]; if (nl > 0 && keep[nl]) mine = true;
                }
            }
            if (!mine) continue;
            int i = gy * stride + gx * 4, j = (y * cw + x) * 4, d = D(p, i, r, g, b);
            int a = d >= 70 ? 255 : (d <= 8 ? 0 : (d - 8) * 255 / 62);
            if (a > 0) {
                int[] bgc = { b, g, r };
                for (int k = 0; k < 3; k++) { int v = bgc[k] + (p[i + k] - bgc[k]) * 255 / a; o[j + k] = (byte)Math.Max(0, Math.Min(255, v)); }
                o[j + 3] = (byte)a;
            }
            p[i] = (byte)b; p[i + 1] = (byte)g; p[i + 2] = (byte)r; p[i + 3] = 255;
        }
        return o;
    }
}
'@

# ---------------------------------------------------------------- 1. OCR (built into Windows)
Write-Host 'Reading the text in the picture...'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op, [type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(-1) | Out-Null; $task.Result }
$asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
function AwaitAction($op) { $asTaskAction.Invoke($null, @($op)).Wait(-1) | Out-Null }

# ---------------------------------------------------------------- 0. accept PDF / WEBP / GIF / HEIC ... -> one PNG
# System.Drawing only reads PNG / JPG / BMP well ("Out of memory" for other files), so other files are converted first
$origImage = $Image; $tempPng = $null
function WinRtStreamToBitmap($ras) {
    $ras.Seek(0); $ms = New-Object System.IO.MemoryStream
    [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($ras.GetInputStreamAt(0)).CopyTo($ms)
    $ms.Position = 0; [System.Drawing.Bitmap]::FromStream($ms)
}
$head = New-Object byte[] 5; $fs = [IO.File]::OpenRead($Image); [void]$fs.Read($head, 0, 5); $fs.Close()
$isPdf = [Text.Encoding]::ASCII.GetString($head) -eq '%PDF-'
$src = $null
if (-not $isPdf) { try { $src = [System.Drawing.Bitmap]::FromFile($Image) } catch { $src = $null } }
if (-not $src) {
    $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    $sf = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Image)) ([Windows.Storage.StorageFile])
    if ($isPdf) {
        Write-Host 'PDF found - turning its pages into one picture...'
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $pdf = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)) ([Windows.Data.Pdf.PdfDocument])
        $pages = New-Object System.Collections.Generic.List[object]
        for ($pi = 0; $pi -lt $pdf.PageCount; $pi++) {
            $pg = $pdf.GetPage($pi); $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions; $opt.DestinationWidth = 1800
            $ras = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            AwaitAction ($pg.RenderToStreamAsync($ras, $opt)); $pages.Add((WinRtStreamToBitmap $ras)); $pg.Dispose()
        }
        Write-Host "  $($pages.Count) page(s)"
        $tw = ($pages | Measure-Object Width -Maximum).Maximum; $th = ($pages | Measure-Object Height -Sum).Sum
        $src = New-Object System.Drawing.Bitmap ([int]$tw), ([int]$th)
        $g0 = [System.Drawing.Graphics]::FromImage($src); $g0.Clear([System.Drawing.Color]::White); $yy = 0
        foreach ($pgb in $pages) { $g0.DrawImage($pgb, 0, $yy, $pgb.Width, $pgb.Height); $yy += $pgb.Height; $pgb.Dispose() }
        $g0.Dispose()
    } else {
        Write-Host 'Converting the picture to PNG first...'
        $st0 = Await ($sf.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        try { $d0 = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($st0)) ([Windows.Graphics.Imaging.BitmapDecoder]) }
        catch { throw "This file type cannot be read: $([IO.Path]::GetFileName($Image)). Please save it as PNG or JPG (in Canva: Share > Download > PNG)." }
        $sb0 = Await ($d0.GetSoftwareBitmapAsync([Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8, [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied)) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $out0 = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
        $e0 = Await ([Windows.Graphics.Imaging.BitmapEncoder]::CreateAsync([Windows.Graphics.Imaging.BitmapEncoder]::PngEncoderId, $out0)) ([Windows.Graphics.Imaging.BitmapEncoder])
        $e0.SetSoftwareBitmap($sb0); AwaitAction ($e0.FlushAsync()); $st0.Dispose()
        $src = WinRtStreamToBitmap $out0
    }
    $tempPng = Join-Path $env:TEMP "design_$PID.png"; $src.Save($tempPng, [System.Drawing.Imaging.ImageFormat]::Png)
    $src.Dispose(); $src = [System.Drawing.Bitmap]::FromFile($tempPng); $Image = $tempPng
}
$W = $src.Width; $H = $src.Height
$ocrPath = $Image; $ocrScale = 1.0
if ([Math]::Max($W, $H) -gt 9000) {   # OCR limit is 10000px
    $ocrScale = 9000 / [Math]::Max($W, $H)
    $tmp = New-Object System.Drawing.Bitmap $src, ([int]($W * $ocrScale)), ([int]($H * $ocrScale))
    $ocrPath = Join-Path $env:TEMP "ocr_$PID.png"; $tmp.Save($ocrPath, [System.Drawing.Imaging.ImageFormat]::Png); $tmp.Dispose()
}
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($ocrPath)) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$sb = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$eng = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if (-not $eng) { throw 'Windows OCR is not available for your language. Add an English language pack in Windows Settings > Language.' }
$res = Await ($eng.RecognizeAsync($sb)) ([Windows.Media.Ocr.OcrResult])
$stream.Dispose()
if ($ocrPath -ne $Image) { Remove-Item $ocrPath -ErrorAction SilentlyContinue }

$lines = foreach ($ln in $res.Lines) {
    $r = @($ln.Words | ForEach-Object { $_.BoundingRect })
    if (-not $r.Count) { continue }
    $x0 = ($r | ForEach-Object { $_.X } | Measure-Object -Minimum).Minimum / $ocrScale
    $y0 = ($r | ForEach-Object { $_.Y } | Measure-Object -Minimum).Minimum / $ocrScale
    $x1 = ($r | ForEach-Object { $_.X + $_.Width } | Measure-Object -Maximum).Maximum / $ocrScale
    $y1 = ($r | ForEach-Object { $_.Y + $_.Height } | Measure-Object -Maximum).Maximum / $ocrScale
    # OCR boxes hug the letters: without descenders (g, p, y, comma...) the box is only the cap height (~72% of the font size)
    $hasDesc = $ln.Text -cmatch '[gjpqyQJ,;()\[\]{}|@$]'
    [pscustomobject]@{ text = $ln.Text; x0 = $x0; y0 = $y0; x1 = $x1; y1 = $y1; h = $y1 - $y0; fs = ($y1 - $y0) * $(if ($hasDesc) { 1.08 } else { 1.38 }); caps = ($ln.Text -cnotmatch '[a-z]') }
}
$lines = @($lines | Sort-Object y0, x0)
Write-Host "Found $($lines.Count) lines of text."

# ---------------------------------------------------------------- 2. colours (fast pixel access)
$rect = New-Object System.Drawing.Rectangle 0, 0, $W, $H
$work = $src.Clone($rect, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb); $src.Dispose()
$bd = $work.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadWrite, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$stride = $bd.Stride; $px = New-Object byte[] ($stride * $H)
[Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $px, 0, $px.Length)
function Px([int]$x, [int]$y) { $x = [Math]::Min($W - 1, [Math]::Max(0, $x)); $y = [Math]::Min($H - 1, [Math]::Max(0, $y)); $i = $y * $stride + $x * 4; , @($px[$i + 2], $px[$i + 1], $px[$i]) }
function Median($a) { $s = @($a | Sort-Object); $s[[int]($s.Count / 2)] }
function Hex($c) { '#{0:X2}{1:X2}{2:X2}' -f [int]$c[0], [int]$c[1], [int]$c[2] }
function Dist($a, $b) { [Math]::Sqrt(($a[0] - $b[0]) * ($a[0] - $b[0]) + ($a[1] - $b[1]) * ($a[1] - $b[1]) + ($a[2] - $b[2]) * ($a[2] - $b[2])) }

# justified text: OCR splits a line where the word gap is wide -> join pieces on the same row when the gap between them is empty
function GapIsEmpty($a, $b) {
    $gy0 = [int][Math]::Max($a.y0, $b.y0); $gy1 = [int][Math]::Min($a.y1, $b.y1); if ($gy1 -le $gy0) { $gy0 = [int]$a.y0; $gy1 = [int]$a.y1 }
    [DesignOps]::GapEmpty($px, $stride, $W, $H, [int]$a.x1 + 2, $gy0, [int]$b.x0 - 2, $gy1, 60)
}
$joined = New-Object System.Collections.Generic.List[object]
foreach ($l in ($lines | Sort-Object x0)) {
    $m = $null
    foreach ($c in $joined) {
        # same row: the two pieces overlap vertically for most of the smaller one (letters with/without ascenders differ in height)
        $ov = [Math]::Min($c.y1, $l.y1) - [Math]::Max($c.y0, $l.y0)
        $gap = $l.x0 - $c.x1
        $sameRow = $ov -ge [Math]::Min($c.h, $l.h) * 0.6 -and [Math]::Abs(($c.y0 + $c.y1) - ($l.y0 + $l.y1)) / 2 -lt [Math]::Max($c.h, $l.h) * 0.45
        $hh = [Math]::Max($c.h, $l.h)
        # a very wide gap is a justified word gap when the rows just above/below have text crossing it;
        # a column gutter stays empty above and below
        $okGap = $gap -lt $hh * 3
        if (-not $okGap -and $gap -lt $W * 0.6) {
            $gx0 = [int]$c.x1 + 2; $gx1 = [int]$l.x0 - 2
            $above = -not [DesignOps]::GapEmpty($px, $stride, $W, $H, $gx0, [int]([Math]::Min($c.y0, $l.y0) - $hh * 2.2), $gx1, [int]([Math]::Min($c.y0, $l.y0) - $hh * 0.4), 60)
            $below = -not [DesignOps]::GapEmpty($px, $stride, $W, $H, $gx0, [int]([Math]::Max($c.y1, $l.y1) + $hh * 0.4), $gx1, [int]([Math]::Max($c.y1, $l.y1) + $hh * 2.2), 60)
            $okGap = $above -or $below
        }
        if ($sameRow -and $gap -gt -3 -and $okGap -and (GapIsEmpty $c $l)) { $m = $c; break }
    }
    if ($m) {
        $m.text = $m.text + ' ' + $l.text; $m.x1 = [Math]::Max($m.x1, $l.x1); $m.y0 = [Math]::Min($m.y0, $l.y0); $m.y1 = [Math]::Max($m.y1, $l.y1)
        $m.h = $m.y1 - $m.y0; $m.fs = [Math]::Max($m.fs, $l.fs); $m.caps = $m.caps -and $l.caps
    } else { $joined.Add($l) }
}
if ($joined.Count -lt $lines.Count) { Write-Host "Joined split lines: $($lines.Count) -> $($joined.Count)." }
$lines = @($joined | Sort-Object y0, x0)

foreach ($l in $lines) {
    $pad = [Math]::Max(3, [int]($l.h * 0.18))
    # background = median colour around the line; text colour = the pixels that differ most; ink ratio -> bold guess
    $st = [DesignOps]::LineStats($px, $stride, $W, $H, [int]$l.x0, [int]$l.y0, [int]$l.x1, [int]$l.y1, $pad)
    $bg = @($st[0], $st[1], $st[2]); $fg = @($st[3], $st[4], $st[5]); $inkRatio = $st[6]
    # ink density, normalised for lines whose box is only the cap height -> bold if the letters are "heavy"
    $inkN = $inkRatio * ($l.h * 1.08 / [Math]::Max(1, $l.fs))
    $l | Add-Member bg $bg; $l | Add-Member fg $fg; $l | Add-Member bold ($inkN -gt 0.265); $l | Add-Member pad $pad; $l | Add-Member busy $st[7]
}
# text printed on top of a photo (busy surroundings) stays part of the picture - painting it out would damage the photo
$onPhoto = @($lines | Where-Object { $_.busy -gt 30 })
if ($onPhoto.Count) { Write-Host "Kept $($onPhoto.Count) line(s) that sit on photos as part of the picture." }
$lines = @($lines | Where-Object { $_.busy -le 30 })
# bullets (dots / icons just left of a line) that OCR did not read: include them in the text box
foreach ($l in $lines) {
    $bx0 = [int]($l.x0 - $l.h * 1.6); $bx1 = [int]($l.x0 - 2)
    if ($bx0 -lt 0 -or $l.text -match '^[\u2022\-\*o]\s') { continue }
    $share = [DesignOps]::InkShare($px, $stride, $W, $H, $bx0, [int]$l.y0, $bx1, [int]$l.y1, [int]$l.bg[0], [int]$l.bg[1], [int]$l.bg[2], 60)
    $edge = [DesignOps]::InkShare($px, $stride, $W, $H, $bx0, [int]$l.y0, $bx0 + 2, [int]$l.y1, [int]$l.bg[0], [int]$l.bg[1], [int]$l.bg[2], 60)
    if ($share -gt 0.03 -and $share -lt 0.4 -and $edge -lt 0.05) { $l.x0 = $bx0 + 2; $l.text = "$([char]0x2022) " + $l.text }
}

# ---------------------------------------------------------------- 3. paint the text out of the picture
$k = $DESIGN_W / $W
foreach ($l in $lines) {
    $x0 = [Math]::Max(0, [int]($l.x0 - $l.pad)); $y0 = [Math]::Max(0, [int]($l.y0 - $l.pad))
    $x1 = [Math]::Min($W - 1, [int]($l.x1 + $l.pad)); $y1 = [Math]::Min($H - 1, [int]($l.y1 + $l.pad))
    [DesignOps]::Fill($px, $stride, $x0, $y0, $x1, $y1, [int]$l.bg[0], [int]$l.bg[1], [int]$l.bg[2])
}

# ---------------------------------------------------------------- 3b. pictures / icons -> separate movable picture boxes
Write-Host 'Finding pictures and icons...'
$border = New-Object System.Collections.Generic.List[object]
$bs = [Math]::Max(1, [int]([Math]::Max($W, $H) / 400))
for ($x = 0; $x -lt $W; $x += $bs) { $border.Add((Px $x 2)); $border.Add((Px $x ($H - 3))) }
for ($y = 0; $y -lt $H; $y += $bs) { $border.Add((Px 2 $y)); $border.Add((Px ($W - 3) $y)) }
$page = @((Median ($border | ForEach-Object { $_[0] })), (Median ($border | ForEach-Object { $_[1] })), (Median ($border | ForEach-Object { $_[2] })))
$comps = New-Object 'System.Collections.Generic.List[int[]]'
$lab = [DesignOps]::Label($px, $stride, $W, $H, [int]$page[0], [int]$page[1], [int]$page[2], 40, $comps)
$s = $W / 1000
$cand = @(foreach ($c in $comps) {
    $cw = $c[3] - $c[1] + 1; $ch = $c[4] - $c[2] + 1; $fill = $c[5] / ($cw * $ch)
    # keep picture-like areas: not tiny (dashes, dots), not page-sized, not solid boxes/bars, not thin outlines
    $solidBar = $fill -ge 0.86 -and ($cw -gt 120 * $s -or $ch -gt 120 * $s)   # big filled boxes / bars stay in the design
    if ([Math]::Min($cw, $ch) -ge 14 * $s -and $cw -le $W * 0.45 -and $ch -le $H * 0.4 -and -not $solidBar -and $fill -gt 0.10) {
        $ids = New-Object 'System.Collections.Generic.List[int]'; $ids.Add([int]$c[0])
        [pscustomobject]@{ ids = $ids; x0 = $c[1]; y0 = $c[2]; x1 = $c[3]; y1 = $c[4] }
    }
})
# parts that touch or almost touch belong to the same picture
$gapPx = 8 * $s
do {
    $changed = $false
    for ($ia = 0; $ia -lt $cand.Count; $ia++) { for ($ib = $cand.Count - 1; $ib -gt $ia; $ib--) {
        $ga = $cand[$ia]; $gb = $cand[$ib]
        if ($gb.x0 -le $ga.x1 + $gapPx -and $ga.x0 -le $gb.x1 + $gapPx -and $gb.y0 -le $ga.y1 + $gapPx -and $ga.y0 -le $gb.y1 + $gapPx) {
            $ga.ids.AddRange($gb.ids); $ga.x0 = [Math]::Min($ga.x0, $gb.x0); $ga.y0 = [Math]::Min($ga.y0, $gb.y0); $ga.x1 = [Math]::Max($ga.x1, $gb.x1); $ga.y1 = [Math]::Max($ga.y1, $gb.y1)
            $cand = @($cand | Where-Object { $_ -ne $gb }); $changed = $true
        }
    } }
} while ($changed)
$gfx = New-Object System.Collections.Generic.List[object]
$keep = New-Object bool[] ($comps.Count + 2)
foreach ($g in $cand) {
    foreach ($id in $g.ids) { $keep[$id] = $true }
    $e = [int][Math]::Max(1, 2 * $s)
    $x0 = [Math]::Max(0, $g.x0 - $e); $y0 = [Math]::Max(0, $g.y0 - $e); $x1 = [Math]::Min($W - 1, $g.x1 + $e); $y1 = [Math]::Min($H - 1, $g.y1 + $e)
    $bytes = [DesignOps]::Extract($px, $stride, $lab, $W, $keep, $x0, $y0, $x1, $y1, [int]$page[0], [int]$page[1], [int]$page[2])
    foreach ($id in $g.ids) { $keep[$id] = $false }
    $cw = $x1 - $x0 + 1; $ch = $y1 - $y0 + 1
    $bm = New-Object System.Drawing.Bitmap $cw, $ch, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $bl = $bm.LockBits((New-Object System.Drawing.Rectangle 0, 0, $cw, $ch), [System.Drawing.Imaging.ImageLockMode]::WriteOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    for ($row = 0; $row -lt $ch; $row++) { [Runtime.InteropServices.Marshal]::Copy($bytes, $row * $cw * 4, [IntPtr]($bl.Scan0.ToInt64() + $row * $bl.Stride), $cw * 4) }
    $bm.UnlockBits($bl)
    $outScale = [Math]::Min(1, 600 / [Math]::Max($cw, $ch))
    if ($outScale -lt 1) { $sm = New-Object System.Drawing.Bitmap $bm, ([int]($cw * $outScale)), ([int]($ch * $outScale)); $bm.Dispose(); $bm = $sm }
    $ms2 = New-Object System.IO.MemoryStream; $bm.Save($ms2, [System.Drawing.Imaging.ImageFormat]::Png); $bm.Dispose()
    $gfx.Add([pscustomobject]@{ x0 = $x0; y0 = $y0; w = $cw; h = $ch; url = 'data:image/png;base64,' + [Convert]::ToBase64String($ms2.ToArray()) })
}
Write-Host "Found $($gfx.Count) pictures/icons."
[Runtime.InteropServices.Marshal]::Copy($px, 0, $bd.Scan0, $px.Length)
$work.UnlockBits($bd)

# ---------------------------------------------------------------- 4. group lines into paragraphs -> text boxes
$groups = New-Object System.Collections.Generic.List[object]
foreach ($l in $lines) {
    $g = $null
    foreach ($cand in $groups) {
        $last = $cand[$cand.Count - 1]
        $sameCol = [Math]::Abs($l.x0 - $cand[0].x0) -lt [Math]::Max(8, $l.h * 0.8) -or ($l.x0 -ge $cand[0].x0 - 4 -and $l.x1 -le ($cand | Measure-Object x1 -Maximum).Maximum + 4)
        $gap = $l.y0 - $last.y1
        $ratio = $l.fs / [Math]::Max(1, $last.fs)
        if ($sameCol -and $gap -ge -2 -and $gap -lt $last.fs * 0.9 -and $ratio -gt 0.75 -and $ratio -lt 1.33 -and (Dist $l.fg $last.fg) -lt 90) { $g = $cand; break }
    }
    if ($g) { $g.Add($l) } else { $n = New-Object System.Collections.Generic.List[object]; $n.Add($l); $groups.Add($n) }
}
$items = New-Object System.Collections.Generic.List[object]
$z = 1
foreach ($gp in $gfx) {   # pictures first, so text stays on top
    $items.Add([ordered]@{ id = 'g' + $z; type = 'image'; gfx = $true; x = [Math]::Round($gp.x0 * $k); y = [Math]::Round($gp.y0 * $k)
        w = [Math]::Max(8, [Math]::Round($gp.w * $k)); h = [Math]::Max(8, [Math]::Round($gp.h * $k)); z = $z; img = $gp.url; round = $false })
    $z++
}
foreach ($g in $groups) {
    $f = $g[0]
    $avgH = ($g | Measure-Object h -Average).Average
    $size = [Math]::Max(8, [Math]::Round(($g | Measure-Object fs -Average).Average * $k))
    $lh = if ($g.Count -gt 1) { [Math]::Round(((($g[$g.Count - 1].y0 - $f.y0) / ($g.Count - 1)) * $k) / $size, 2) } else { 1.2 }
    $lh = [Math]::Min(2.2, [Math]::Max(1.0, $lh))
    $gx0 = ($g | Measure-Object x0 -Minimum).Minimum; $gx1 = ($g | Measure-Object x1 -Maximum).Maximum
    $lineBox = $lh * $size
    $x = [Math]::Round($gx0 * $k - 5)
    # put the top of the letters where OCR found them (Arial: content area 1.117em, cap top 0.189em below it; 2px box padding)
    $y = [Math]::Round($f.y0 * $k - (($lineBox - 1.117 * $size) / 2 + 0.189 * $size) - 2)
    $boxW = [Math]::Round(($gx1 - $gx0) * $k * 1.06 + 14)
    $boxH = [Math]::Round($g.Count * $lineBox + 8)
    $items.Add([ordered]@{
        id = 't' + $z; type = 'text'; x = $x; y = $y; w = $boxW; h = $boxH; z = $z
        text = (($g | ForEach-Object { $_.text }) -join "`n"); size = $size; lh = $lh
        color = (Hex $f.fg); bold = [bool]$f.bold; italic = $false; align = 'left'; bg = ''; font = 'Arial'
    })
    $z++
}

# ---------------------------------------------------------------- 5. save: cleaned picture (JPEG, max 1800px wide) + boxes
$outW = [Math]::Min(1800, $W); $outH = [int][Math]::Round($H * $outW / $W)
$small = New-Object System.Drawing.Bitmap $outW, $outH
$gr = [System.Drawing.Graphics]::FromImage($small); $gr.InterpolationMode = 'HighQualityBicubic'; $gr.DrawImage($work, 0, 0, $outW, $outH); $gr.Dispose(); $work.Dispose()
$ms = New-Object System.IO.MemoryStream
$enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
$ep = New-Object System.Drawing.Imaging.EncoderParameters 1
$ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), ([long]88)
$small.Save($ms, $enc, $ep); $small.Dispose()
$data = [ordered]@{
    type = 'projecthub-design'; version = 1; source = [IO.Path]::GetFileName($origImage)
    canvas = [ordered]@{ bg = 'data:image/jpeg;base64,' + [Convert]::ToBase64String($ms.ToArray()); h = [int][Math]::Round($DESIGN_W * $H / $W); items = $items.ToArray() }
}
[IO.File]::WriteAllText($OutFile, ($data | ConvertTo-Json -Depth 6 -Compress), (New-Object System.Text.UTF8Encoding $false))
Write-Host ''
Write-Host "Done: $($items.Count - $gfx.Count) editable text boxes and $($gfx.Count) movable pictures."
Write-Host "Saved: $OutFile"
if ($tempPng) { Remove-Item $tempPng -ErrorAction SilentlyContinue }
Write-Host 'Next: open Project_Hub.html > Newsletter > Upload new format, and choose this .json file.'
