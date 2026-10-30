param([string]$In, [string]$Out, [int]$Pad = 4)
# crops the white / transparent border around a logo so it can be shown larger
Add-Type -AssemblyName System.Drawing
$src = [System.Drawing.Bitmap]::FromFile($In)
$W = $src.Width; $H = $src.Height
$minX = $W; $minY = $H; $maxX = -1; $maxY = -1
for ($y = 0; $y -lt $H; $y++) {
    for ($x = 0; $x -lt $W; $x++) {
        $c = $src.GetPixel($x, $y)
        if ($c.A -gt 20 -and ($c.R -lt 235 -or $c.G -lt 235 -or $c.B -lt 235)) {
            if ($x -lt $minX) { $minX = $x }; if ($x -gt $maxX) { $maxX = $x }
            if ($y -lt $minY) { $minY = $y }; if ($y -gt $maxY) { $maxY = $y }
        }
    }
}
$minX = [Math]::Max(0, $minX - $Pad); $minY = [Math]::Max(0, $minY - $Pad)
$maxX = [Math]::Min($W - 1, $maxX + $Pad); $maxY = [Math]::Min($H - 1, $maxY + $Pad)
$rect = New-Object System.Drawing.Rectangle $minX, $minY, ($maxX - $minX + 1), ($maxY - $minY + 1)
$crop = $src.Clone($rect, $src.PixelFormat)
$src.Dispose()
$crop.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png); $crop.Dispose()
"{0}: {1}x{2} -> {3}x{4}" -f [IO.Path]::GetFileName($Out), $W, $H, $rect.Width, $rect.Height
