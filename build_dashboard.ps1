param(
    [string]$OutFile = "$PSScriptRoot\Project_Consolidation_Dashboard.xlsx",
    # existing workbook whose Projects table data is carried over (keeps projects users have added)
    [string]$DataFrom = "$PSScriptRoot\Project_Consolidation_Dashboard.xlsx"
)

$ErrorActionPreference = 'Stop'
$M = [System.Reflection.Missing]::Value
function RGB($r, $g, $b) { [int]($r + 256 * $g + 65536 * $b) }

$NAVY = RGB 31 56 100; $BLUE = RGB 46 117 182; $GREEN = RGB 84 130 53; $ORANGE = RGB 237 125 49
$RED = RGB 192 0 0; $PURPLE = RGB 112 48 160; $TEAL = RGB 0 128 128; $WHITE = RGB 255 255 255
$LIGHT = RGB 242 242 242; $GREY = RGB 166 166 166
function HEX($h) { RGB ([Convert]::ToInt32($h.Substring(1, 2), 16)) ([Convert]::ToInt32($h.Substring(3, 2), 16)) ([Convert]::ToInt32($h.Substring(5, 2), 16)) }
$P1 = HEX '#C693FF'; $P2 = HEX '#C3BBFD'; $P3 = HEX '#A1DCF9'; $P4 = HEX '#C5FFFF'
$INK = HEX '#2E2257'; $SUB = HEX '#5E5480'; $OVERDUE = HEX '#FFB3BE'; $OVERDUE_INK = HEX '#8B1E2F'
$NAVBG = $P2; $NAVW = 22
$TRACK = HEX '#ECE8FB'   # empty part of the gauge / progress bar on white cards
# vivid accents used on top of the pastel backgrounds (charts, icons, status, bars)
$A_VIOLET = HEX '#7C3AED'; $A_AMBER = HEX '#F59E0B'; $A_EMERALD = HEX '#10B981'
$A_ROSE = HEX '#F43F5E'; $A_SKY = HEX '#0EA5E9'; $A_TEAL = HEX '#14B8A6'
$STC = @{
    'Completed' = @((HEX '#D1FAE5'), (HEX '#065F46')); 'Ongoing' = @((HEX '#FEF3C7'), (HEX '#92400E'))
    'Overdue'   = @((HEX '#FFE4E6'), (HEX '#9F1239')); 'Due Soon' = @((HEX '#FEF3C7'), (HEX '#92400E'))
    'On Track'  = @((HEX '#E0F2FE'), (HEX '#075985')); 'High' = @((HEX '#FFE4E6'), (HEX '#9F1239'))
}

# Left sidebar navigation (shape buttons with internal hyperlinks - no macros needed)
function Add-Nav($ws, $active) {
    $ws.Columns.Item(1).ColumnWidth = $NAVW
    $ws.Columns.Item(1).Interior.Color = $NAVBG
    $ws.Columns.Item(1).Borders.Item(10).Color = $P1
    $L = 6; $W = $ws.Columns.Item(1).Width - 12; $T = 10
    $hd = $ws.Shapes.AddTextbox(1, $L + 4, $T, $W, 20)
    $hd.Fill.Visible = 0; $hd.Line.Visible = 0
    $hd.TextFrame2.TextRange.Text = 'PROJECT HUB'
    $hd.TextFrame2.TextRange.Font.Name = 'Segoe UI'; $hd.TextFrame2.TextRange.Font.Size = 9
    $hd.TextFrame2.TextRange.Font.Bold = 1; $hd.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $SUB
    $hd.Placement = 3
    $T += 26
    foreach ($it in @(@('Dashboard', 0xE80F), @('Projects', 0xE716), @('Timeline', 0xE787), @('Newsletter', 0xE715))) {
        $on = $it[0] -eq $active
        $b = $ws.Shapes.AddShape(5, $L, $T, $W, 30)
        $b.Name = "nav_$($it[0])"
        $b.Adjustments.Item(1) = 0.2
        $b.Line.Visible = 0
        $b.Fill.ForeColor.RGB = $(if ($on) { $A_VIOLET } else { $NAVBG })
        if ($on) { $b.Shadow.Visible = -1; $b.Shadow.Transparency = 0.7; $b.Shadow.OffsetX = 1; $b.Shadow.OffsetY = 2 }
        $tf = $b.TextFrame2; $tf.MarginLeft = 10; $tf.VerticalAnchor = 3; $tf.WordWrap = 0
        $tr = $tf.TextRange
        $tr.Text = [string][char]$it[1] + '   ' + $it[0]
        $tr.ParagraphFormat.Alignment = 1
        $tr.Font.Name = 'Segoe UI'; $tr.Font.Size = 11; $tr.Font.Bold = 1; $tr.Font.Fill.ForeColor.RGB = $(if ($on) { $WHITE } else { $INK })
        $ic = $tr.Characters(1, 1); $ic.Font.Name = 'Segoe MDL2 Assets'; $ic.Font.Size = 14; $ic.Font.Bold = 0
        $b.Placement = 3
        $ws.Hyperlinks.Add($b, '', "'$($it[0])'!A1", "Go to $($it[0])") | Out-Null
        $T += 34
    }
}

# Soft abstract background (lavender top-left, near-white centre, blue top-right, mint bottom-right,
# translucent waves, dot grids, thin lines) drawn with System.Drawing and saved as a JPEG.
function New-BackgroundImage([string]$path, [int]$W, [int]$H) {
    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object System.Drawing.Bitmap $W, $H
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $C = { param($a, $hex) [System.Drawing.Color]::FromArgb($a, [Convert]::ToInt32($hex.Substring(1, 2), 16), [Convert]::ToInt32($hex.Substring(3, 2), 16), [Convert]::ToInt32($hex.Substring(5, 2), 16)) }
    $S = [double]$W
    $g.Clear((& $C 255 '#F3F3F9'))
    function Blob($cx, $cy, $r, $a, $hex) {
        $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
        $gp.AddEllipse([single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r))
        $br = New-Object System.Drawing.Drawing2D.PathGradientBrush $gp
        $br.CenterColor = & $C $a $hex; $br.SurroundColors = @(& $C 0 $hex)
        $g.FillPath($br, $gp); $br.Dispose(); $gp.Dispose()
    }
    Blob 0 0 (0.80 * $S) 255 '#C693FF'
    Blob (0.10 * $S) (0.08 * $S) (0.50 * $S) 210 '#C3BBFD'
    Blob 0 (0.45 * $H) (0.50 * $S) 130 '#C3BBFD'
    Blob $S 0 (0.62 * $S) 255 '#A1DCF9'
    Blob $S $H (0.60 * $S) 255 '#A6EFE2'
    Blob ($S * 0.95) ($H * 0.97) (0.30 * $S) 230 '#C5FFFF'
    Blob 0 $H (0.50 * $S) 150 '#A1DCF9'
    Blob $S ($H * 0.5) (0.40 * $S) 110 '#A1DCF9'
    $P = { param($x, $y) New-Object System.Drawing.PointF ([single]$x), ([single]$y) }
    # translucent white overlays with wavy edges -> layered wave look
    $waves = @(
        @((& $P 0 (0.30 * $S)), (& $P (0.12 * $S) (0.20 * $S)), (& $P (0.22 * $S) (0.20 * $S)), (& $P (0.30 * $S) (0.10 * $S)), (& $P (0.36 * $S) 0)),
        @((& $P (0.66 * $S) 0), (& $P (0.70 * $S) (0.10 * $S)), (& $P (0.82 * $S) (0.07 * $S)), (& $P (0.88 * $S) (0.14 * $S)), (& $P $S (0.24 * $S)))
    )
    foreach ($wv in $waves) {
        $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
        $gp.AddBeziers([System.Drawing.PointF[]]@($wv[0], $wv[1], $wv[2], $wv[3]))
        $gp.AddLine($wv[3], $wv[4])
        if ($wv[0].X -eq 0) { $gp.AddLines([System.Drawing.PointF[]]@((& $P $S 0), (& $P $S $H), (& $P 0 $H))) }
        else { $gp.AddLines([System.Drawing.PointF[]]@((& $P $S $H), (& $P 0 $H), (& $P 0 0))) }
        $gp.CloseFigure()
        $br = New-Object System.Drawing.SolidBrush (& $C 70 '#FFFFFF'); $g.FillPath($br, $gp); $br.Dispose(); $gp.Dispose()
    }
    $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
    $gp.AddBezier((& $P ($S * 0.80) $H), (& $P ($S * 0.86) ($H - 0.12 * $S)), (& $P ($S * 0.94) ($H - 0.10 * $S)), (& $P $S ($H - 0.22 * $S)))
    $gp.AddLines([System.Drawing.PointF[]]@((& $P $S 0), (& $P 0 0), (& $P 0 $H))); $gp.CloseFigure()
    $br = New-Object System.Drawing.SolidBrush (& $C 55 '#FFFFFF'); $g.FillPath($br, $gp); $br.Dispose(); $gp.Dispose()
    # contour lines near the top-left wave
    $pen = New-Object System.Drawing.Pen (& $C 120 '#FFFFFF'), ([single]([Math]::Max(1, $S / 1400)))
    for ($i = 1; $i -le 7; $i++) {
        $o = $i * $S * 0.008
        $g.DrawBezier($pen, (& $P 0 (0.24 * $S - $o)), (& $P (0.12 * $S) (0.14 * $S - $o)), (& $P (0.22 * $S) (0.14 * $S - $o)), (& $P (0.28 * $S - $o / 2) (0.0)))
    }
    # dot grids
    $dot = New-Object System.Drawing.SolidBrush (& $C 210 '#FFFFFF'); $dr = $S * 0.0026
    for ($gr = 0; $gr -lt 6; $gr++) { for ($gc = 0; $gc -lt 6; $gc++) { $g.FillEllipse($dot, [single](0.015 * $S + $gc * 0.017 * $S), [single](0.02 * $S + $gr * 0.017 * $S), [single](2 * $dr), [single](2 * $dr)) } }
    for ($gr = 0; $gr -lt 10; $gr++) { for ($gc = 0; $gc -lt 5; $gc++) { $g.FillEllipse($dot, [single](0.915 * $S + $gc * 0.016 * $S), [single](0.06 * $S + $gr * 0.016 * $S), [single](2 * $dr), [single](2 * $dr)) } }
    for ($gr = 0; $gr -lt 5; $gr++) { for ($gc = 0; $gc -lt 6; $gc++) { $g.FillEllipse($dot, [single](0.03 * $S + $gc * 0.017 * $S), [single]($H - 0.10 * $S + $gr * 0.017 * $S), [single](2 * $dr), [single](2 * $dr)) } }
    # thin diagonal lines and outline circles
    $lp = New-Object System.Drawing.Pen (& $C 150 '#FFFFFF'), ([single]([Math]::Max(1, $S / 1100)))
    foreach ($ln in @(@(0.51, 0.11, 0.59, 0.04), @(0.52, 0.13, 0.60, 0.06), @(0.66, 0.20, 0.76, 0.10), @(0.00, 0.34, 0.04, 0.30), @(0.00, 0.37, 0.05, 0.32))) {
        $g.DrawLine($lp, [single]($ln[0] * $S), [single]($ln[1] * $S), [single]($ln[2] * $S), [single]($ln[3] * $S))
    }
    foreach ($ln in @(@(0.85, 0.06, 0.97, 0.17), @(0.87, 0.04, 0.98, 0.14), @(0.40, 0.03, 0.47, 0.10))) {
        $g.DrawLine($lp, [single]($ln[0] * $S), [single]($H - $ln[1] * $S), [single]($ln[2] * $S), [single]($H - $ln[3] * $S))
    }
    $g.DrawEllipse($lp, [single](0.645 * $S), [single](0.03 * $S), [single](0.04 * $S), [single](0.04 * $S))
    $g.DrawEllipse($lp, [single](0.92 * $S), [single]($H - 0.06 * $S), [single](0.024 * $S), [single](0.024 * $S))
    $g.DrawEllipse($lp, [single](0.20 * $S), [single](0.40 * $S), [single](0.018 * $S), [single](0.018 * $S))
    # save as high-quality JPEG
    $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
    $ep = New-Object System.Drawing.Imaging.EncoderParameters 1
    $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), ([long]90)
    $bmp.Save($path, $enc, $ep)
    $g.Dispose(); $bmp.Dispose()
}

# ---------------------------------------------------------------- sample data
$cols = 'Project Name', 'POC', 'Developer', 'LOB', 'Status', 'Priority', 'Category', 'Start Date', 'Due Date', 'Deployment Status', 'Progress', 'Link'
$rows = @(
    @('NTE and NOD tracker', 'John Ryan Ariola', 'Shine Abdulhalik, Junie Ann Martinez, Pope Tuazon', 'Payment Integrity', 'Completed', 'High', 'Management Tool', '2026-08-10', '2026-08-21', 'Not Yet Started', 1.00, 'VANGUARD - Power Apps'),
    @('Attendance (with PTO request) tracker', 'Charlon Calugay', 'Justin Umali, Charmaine San Juan', 'Payment Integrity', 'Completed', 'High', 'Workforce Management', '2026-08-10', '2026-08-21', 'Not Yet Started', 1.00, 'Tempus - Power Apps'),
    @('Payment Integrity - Pends Review', 'John Ryan Ariola', 'Shine Abdulhalik, Charlon Calugay', 'Payment Integrity', 'Ongoing', 'Medium', 'Ageing Management', '2026-08-19', '2026-08-26', 'Not Yet Started', 1.00, 'APEX - Power Apps'),
    @('OPIDs Review', 'Justin Umali', 'Justin Umali, Melody Lapena', 'Dispute Intake', 'Completed', 'High', 'Automation', '2026-08-17', '2026-08-21', 'Not Yet Started', 1.00, ''),
    @('Genpro', 'Pope Tuazon', 'Pope Tuazon, Justin Umali', 'Dispute Intake', 'Completed', 'Medium', 'Process Improvement', '2026-08-17', '2026-08-21', 'Submitted', 1.00, ''),
    @('Employee Dashboard', 'John Ryan Ariola', 'John Ryan Ariola', 'Payment Integrity', 'Ongoing', 'High', 'Management Tool', '2026-08-10', '2026-09-04', 'Not Yet Started', 0.50, ''),
    @('New Ansir Apps for NonCOB', 'Shine Abdulhalik', 'Shine Abdulhalik', 'Non-COB', 'Ongoing', 'High', 'Management Tool', '2026-08-10', '2026-09-01', 'Not Yet Started', 0.75, ''),
    @('Ansir Calculator for NonCOB', 'Junie Ann Martinez', 'Junie Ann Martinez', 'Non-COB', 'Ongoing', 'High', 'Management Tool', '2026-08-10', '2026-09-01', 'Not Yet Started', 0.75, ''),
    @('RootCauseAnalysis', 'John Ryan Ariola', 'John Ryan Ariola', 'Payment Integrity', 'Ongoing', 'High', 'Process Improvement', '2026-09-08', '2026-09-11', 'Not Yet Started', 0.75, '')
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    if ($DataFrom -and (Test-Path $DataFrom)) {
        $src = $xl.Workbooks.Open($DataFrom, 0, $true)
        try {
            # the source refreshes its pivots on open; wait until Excel accepts calls again
            $slo = $null
            for ($w = 0; $w -lt 60 -and -not $slo; $w++) {
                try { $slo = $src.Worksheets.Item('Projects').ListObjects.Item('tblProjects') } catch { Start-Sleep 2 }
            }
            $colVals = @{}
            foreach ($cn in $cols) { $colVals[$cn] = $slo.ListColumns.Item($cn).DataBodyRange.Value2 }
            $imported = New-Object System.Collections.Generic.List[object]
            for ($ri = 1; $ri -le $slo.ListRows.Count; $ri++) {
                $one = foreach ($cn in $cols) { $x = $colVals[$cn]; if ($x -is [array]) { $x[$ri, 1] } else { $x } }
                if ("$($one[0])".Trim()) { $imported.Add(@($one)) }
            }
            if ($imported.Count -gt 0) { $rows = $imported.ToArray(); "Imported $($rows.Count) projects from $DataFrom" }
        }
        finally { $src.Close($false) }
    }
    $wb = $xl.Workbooks.Add()
    $xl.Calculation = -4135
    $tcs = $wb.Theme.ThemeColorScheme
    $theme = @{ 3 = $INK; 4 = $P4; 5 = $A_VIOLET; 6 = $A_AMBER; 7 = $A_EMERALD; 8 = $A_SKY; 9 = $A_ROSE; 10 = $A_TEAL }
    foreach ($k in $theme.Keys) { $tcs.Colors($k).RGB = $theme[$k] }
    while ($wb.Worksheets.Count -lt 6) { $wb.Worksheets.Add([System.Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) | Out-Null }
    $dash = $wb.Worksheets.Item(1); $dash.Name = 'Dashboard'
    $proj = $wb.Worksheets.Item(2); $proj.Name = 'Projects'
    $gantt = $wb.Worksheets.Item(3); $gantt.Name = 'Timeline'
    $news = $wb.Worksheets.Item(4); $news.Name = 'Newsletter'
    $pd = $wb.Worksheets.Item(5); $pd.Name = 'PivotData'
    $lists = $wb.Worksheets.Item(6); $lists.Name = 'Lists'

    # ============================================================ PROJECTS
    $proj.Range('A1').Value2 = 'Project Consolidation - Master List'
    $proj.Range('A1').Font.Size = 16; $proj.Range('A1').Font.Bold = $true; $proj.Range('A1').Font.Color = $INK
    $proj.Range('A2').Value2 = 'To add a project, type in the first empty row below the table (grey columns fill in automatically), then click Data > Refresh All to update the Dashboard.'
    $proj.Range('A2').Font.Italic = $true; $proj.Range('A2').Font.Color = $GREY

    $arr = New-Object 'object[,]' ($rows.Count + 1), $cols.Count
    for ($c = 0; $c -lt $cols.Count; $c++) { $arr[0, $c] = $cols[$c] }
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $cols.Count; $c++) {
            $v = $rows[$r][$c]
            if ($c -eq 7 -or $c -eq 8) {
                if ($v -is [string] -and $v) { $v = [datetime]::ParseExact($v, 'yyyy-MM-dd', $null) }
                elseif ($v -is [double]) { $v = [datetime]::FromOADate($v) }
            }
            $arr[($r + 1), $c] = $v
        }
    }
    $last = 3 + $rows.Count
    $proj.Range($proj.Cells(3, 1), $proj.Cells($last, $cols.Count)).Value = $arr
    $lo = $proj.ListObjects.Add(1, $proj.Range($proj.Cells(3, 1), $proj.Cells($last, $cols.Count)), $M, 1)
    $lo.Name = 'tblProjects'; $lo.TableStyle = 'TableStyleMedium2'
    $lo.HeaderRowRange.Font.Color = $WHITE

    $calc = [ordered]@{
        'Duration (Days)' = '=[@[Due Date]]-[@[Start Date]]+1'
        'Days Left'       = '=IF([@Status]="Completed","",[@[Due Date]]-TODAY())'
        'Health'          = '=IF([@Status]="Completed","Completed",IF([@[Due Date]]<TODAY(),"Overdue",IF([@[Due Date]]-TODAY()<=7,"Due Soon","On Track")))'
        'Is Completed'    = '=--([@Status]="Completed")'
        'Is Ongoing'      = '=--([@Status]="Ongoing")'
        'Is Overdue'      = '=--([@Health]="Overdue")'
        'Due (Sort)'      = '=TEXT([@[Due Date]],"yyyy-mm-dd")'
    }
    foreach ($k in $calc.Keys) {
        $lc = $lo.ListColumns.Add(); $lc.Name = $k
        $lc.DataBodyRange.Formula = $calc[$k]
    }
    $lo.ListColumns.Item('Start Date').Range.NumberFormat = 'm/d/yyyy'
    $lo.ListColumns.Item('Due Date').Range.NumberFormat = 'm/d/yyyy'
    $lo.ListColumns.Item('Progress').DataBodyRange.NumberFormat = '0%'
    $lo.ListColumns.Item('Days Left').DataBodyRange.NumberFormat = '0;[Red]-0'
    foreach ($k in $calc.Keys) {
        $cr = $lo.ListColumns.Item($k).DataBodyRange
        $cr.Interior.Color = (HEX '#F1F2F4'); $cr.Font.Color = $SUB
    }

    $db = $lo.ListColumns.Item('Progress').DataBodyRange.FormatConditions.AddDatabar()
    $db.MinPoint.Modify(0, 0); $db.MaxPoint.Modify(0, 1); $db.BarColor.Color = $A_EMERALD

    function Add-TextCF($rng, $text, $fill, $font) {
        $fc = $rng.FormatConditions.Add(1, 3, "=""$text""")
        $fc.Interior.Color = $fill; $fc.Font.Color = $font; $fc.Font.Bold = $true
    }
    $st = $lo.ListColumns.Item('Status').DataBodyRange
    function Add-StatusCF($rng, $text) { Add-TextCF $rng $text $STC[$text][0] $STC[$text][1] }
    Add-StatusCF $st 'Completed'
    Add-StatusCF $st 'Ongoing'
    $hl = $lo.ListColumns.Item('Health').DataBodyRange
    foreach ($h in 'Completed', 'Overdue', 'Due Soon', 'On Track') { Add-StatusCF $hl $h }
    Add-StatusCF $lo.ListColumns.Item('Priority').DataBodyRange 'High'

    $proj.Columns.Item('A:S').AutoFit() | Out-Null
    $proj.Columns.Item('A').ColumnWidth = 36; $proj.Columns.Item('C').ColumnWidth = 44
    $proj.Range('A1').EntireRow.RowHeight = 24
    foreach ($k in 'Is Completed', 'Is Ongoing', 'Is Overdue', 'Due (Sort)') { $lo.ListColumns.Item($k).Range.EntireColumn.Hidden = $true }
    $proj.Columns.Item(1).Insert() | Out-Null
    Add-Nav $proj 'Projects'
    $proj.Activate(); $xl.ActiveWindow.SplitRow = 0; $xl.ActiveWindow.SplitColumn = 2; $xl.ActiveWindow.FreezePanes = $true

    # ============================================================ PIVOTS
    $pc = $wb.PivotCaches().Create(1, 'tblProjects', 5)
    $pc.RefreshOnFileOpen = $true
    function New-Pivot($name, $dest, $rowFields, $colField) {
        $pt = $pc.CreatePivotTable($dest, $name, $true, 5)
        $pos = 1
        foreach ($f in $rowFields) { $pf = $pt.PivotFields($f); $pf.Orientation = 1; $pf.Position = $pos; $pos++ }
        if ($colField) { $pt.PivotFields($colField).Orientation = 2 }
        $pt
    }

    $ptKPI = New-Pivot 'ptKPI' $pd.Range('A3') @() $null
    $ptKPI.AddDataField($ptKPI.PivotFields('Project Name'), 'Total Projects', -4112) | Out-Null
    $ptKPI.AddDataField($ptKPI.PivotFields('Is Completed'), 'Completed Count', -4157) | Out-Null
    $ptKPI.AddDataField($ptKPI.PivotFields('Is Ongoing'), 'Ongoing Count', -4157) | Out-Null
    $ptKPI.AddDataField($ptKPI.PivotFields('Is Overdue'), 'Overdue Count', -4157) | Out-Null
    $ptKPI.AddDataField($ptKPI.PivotFields('Progress'), 'Avg Progress', -4106) | Out-Null

    $ptHealth = New-Pivot 'ptHealth' $pd.Range('I3') @('Health') $null
    $ptHealth.AddDataField($ptHealth.PivotFields('Project Name'), 'Projects', -4112) | Out-Null
    $ptCat = New-Pivot 'ptCategory' $pd.Range('M3') @('Category') $null
    $ptCat.AddDataField($ptCat.PivotFields('Progress'), 'Avg Progress ', -4106).NumberFormat = '0%'
    $ptLOB = New-Pivot 'ptLOB' $pd.Range('Q3') @('LOB') 'Status'
    $ptLOB.AddDataField($ptLOB.PivotFields('Project Name'), 'Projects', -4112) | Out-Null
    $ptPOC = New-Pivot 'ptPOC' $pd.Range('Z3') @('POC') 'Status'
    $ptPOC.AddDataField($ptPOC.PivotFields('Project Name'), 'Projects', -4112) | Out-Null

    # ============================================================ DASHBOARD
    $dash.Activate()
    $xl.ActiveWindow.DisplayGridlines = $false
    $dash.Cells.Font.Name = 'Segoe UI'; $dash.Cells.Font.Size = 10
    $dash.Range('B:Y').ColumnWidth = 9.5
    $w = @{ B = 34; C = 18; D = 16; E = 20; F = 9; G = 10; H = 11; I = 11; J = 11 }
    foreach ($k in $w.Keys) { $dash.Columns.Item($k).ColumnWidth = $w[$k] }
    $dash.Columns.Item('Q').ColumnWidth = 2
    Add-Nav $dash 'Dashboard'
    $L0 = $dash.Range('B1').Left + 10; $W0 = $dash.Range('B1:P1').Width - 10; $GAP = 10
    function Slot($i, $n, $rows) {
        $r = $dash.Range($rows); $sw = ($W0 - $GAP * ($n - 1)) / $n
        @{ L = $L0 + $i * ($sw + $GAP); T = $r.Top + 3; W = $sw; H = $r.Height - 6 }
    }
    function Float($shp) { $shp.Placement = 3 }

    $hdr = $dash.Range('B1:Q2'); $hdr.Interior.Color = $P1; $hdr.Font.Color = $INK
    $dash.Range('B2:Q2').RowHeight = 20; $dash.Range('B2').VerticalAlignment = -4160
    $dash.Range('B3:Q3').Interior.Color = $P3; $dash.Range('3:3').RowHeight = 6
    $dash.Range('B1:B2').IndentLevel = 1
    $dash.Range('B1').Value2 = 'Project Consolidation Dashboard'
    $dash.Range('B1').Font.Size = 20; $dash.Range('B1').Font.Bold = $true
    $dash.Range('1:1').RowHeight = 32
    $dash.Range('B2').Formula = '="As of "&TEXT(TODAY(),"mmmm d, yyyy")&"   |   Use the slicers & timeline below to filter. Edit data on the Projects sheet, then Data > Refresh All."'
    $dash.Range('B2').Font.Size = 9

    # KPI cards
    $kpis = @(
        @('TOTAL PROJECTS', '=IFERROR(GETPIVOTDATA("Total Projects",PivotData!$A$3),0)', '0', $P1, 'in current filter', 0xE8FD, $A_VIOLET),
        @('COMPLETED', '=IFERROR(GETPIVOTDATA("Completed Count",PivotData!$A$3),0)', '0', $P3, 'projects delivered', 0xE73E, $A_EMERALD),
        @('ONGOING', '=IFERROR(GETPIVOTDATA("Ongoing Count",PivotData!$A$3),0)', '0', $P2, 'in progress', 0xE895, $A_AMBER),
        @('OVERDUE', '=IFERROR(GETPIVOTDATA("Overdue Count",PivotData!$A$3),0)', '0', $P4, 'past due, not completed', 0xE7BA, $A_ROSE),
        @('AVG PROGRESS', '=IFERROR(GETPIVOTDATA("Avg Progress",PivotData!$A$3),0)', '0%', $P3, 'average % complete', 0xEAFC, $A_SKY),
        @('COMPLETION RATE', '=IFERROR(GETPIVOTDATA("Completed Count",PivotData!$A$3)/GETPIVOTDATA("Total Projects",PivotData!$A$3),0)', '0%', $P1, 'completed / total', 0xEB05, $A_TEAL)
    )
    function Add-Icon($l, $t, $d, $glyph, $color) {
        $ic = $dash.Shapes.AddShape(9, $l, $t, $d, $d)
        $ic.Fill.ForeColor.RGB = $color; $ic.Line.Visible = 0
        $tf = $ic.TextFrame2; $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
        $tf.VerticalAnchor = 3; $tf.WordWrap = 0
        $tf.TextRange.Text = [string][char]$glyph
        $tf.TextRange.Font.Name = 'Segoe MDL2 Assets'; $tf.TextRange.Font.Size = [int]($d * 0.5)
        $tf.TextRange.Font.Fill.ForeColor.RGB = $WHITE; $tf.TextRange.ParagraphFormat.Alignment = 2
        Float $ic
    }
    $pd.Range('A19').Value2 = 'KPI values (linked to Dashboard cards)'
    function Add-Text([single]$l, [single]$t, [single]$wd, [single]$ht, $text, [single]$size, $bold, [int]$color, $link) {
        $tb = $dash.Shapes.AddTextbox(1, $l, $t, $wd, $ht)
        $tb.Fill.Visible = 0; $tb.Line.Visible = 0
        $tf = $tb.TextFrame2; $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
        if ($link) { $tb.DrawingObject.Formula = $link } else { $tf.TextRange.Text = $text }
        $tf.AutoSize = 0; $tf.WordWrap = 0; $tf.VerticalAnchor = 1; $tb.Top = $t; $tb.Height = $ht
        $tf.TextRange.Font.Size = $size; $tf.TextRange.Font.Bold = [int]$bold; $tf.TextRange.Font.Name = 'Segoe UI'
        $tf.TextRange.Font.Fill.ForeColor.RGB = $color; $tf.TextRange.ParagraphFormat.Alignment = 2
        Float $tb; $tb
    }
    function Clean-Chart($ch) {
        $ch.HasTitle = $false; $ch.HasLegend = $false
        $ch.ChartArea.Format.Fill.Visible = 0; $ch.ChartArea.Format.Line.Visible = 0
        $ch.PlotArea.Format.Fill.Visible = 0
    }
    # circular gauge: 270-degree arc (doughnut with a hidden bottom quarter), % shown in the centre
    function Add-Gauge($l, $t, $d, $row, $color) {
        $pd.Cells($row, 3).Formula = "=MIN(1,MAX(0,B$row))*0.75"
        $pd.Cells($row, 4).Formula = "=0.75-C$row"
        $pd.Cells($row, 5).Value2 = 0.25
        $shp = $dash.Shapes.AddChart2(-1, -4120, $l, $t, $d, $d); Float $shp
        $ch = $shp.Chart
        $ch.SetSourceData($pd.Range("C$($row):E$row"), 1)
        Clean-Chart $ch
        try { $ch.PlotArea.Width = $d; $ch.PlotArea.Height = $d; $ch.PlotArea.Left = 0; $ch.PlotArea.Top = 0 } catch {}
        $ch.ChartGroups(1).DoughnutHoleSize = 68; $ch.ChartGroups(1).FirstSliceAngle = 225
        $s = $ch.SeriesCollection(1); $s.HasDataLabels = $false; $s.Format.Line.Visible = 0
        $s.Points(1).Format.Fill.ForeColor.RGB = $color
        $s.Points(2).Format.Fill.ForeColor.RGB = $TRACK; $s.Points(2).Format.Fill.Transparency = 0
        $s.Points(3).Format.Fill.Visible = 0
        $fs = [int]($d * 0.2); $bw = $d * 0.7
        $null = Add-Text ($l + ($d - $bw) / 2) ($t + $d / 2 - $fs * 0.75) $bw ($fs * 1.5) '' $fs $true $INK "=PivotData!`$B`$$row"
    }
    # horizontal progress bar: 100% stacked bar of value vs remainder
    function Add-ProgressBar($l, $t, $w, $h, $row, $color) {
        $pd.Cells($row, 3).Formula = "=MIN(1,MAX(0,B$row))"
        $pd.Cells($row, 4).Formula = "=1-C$row"
        $pad = 7   # charts keep ~7pt inner padding; grow the shape so the bar itself is $w x $h
        $shp = $dash.Shapes.AddChart2(-1, 59, $l - $pad, $t - $pad, $w + 2 * $pad, $h + 2 * $pad); Float $shp
        $ch = $shp.Chart
        $ch.SetSourceData($pd.Range("C$($row):D$row"), 2)
        Clean-Chart $ch
        foreach ($ax in 1, 2) { try { $ch.Axes($ax).HasMajorGridlines = $false } catch {}; try { $ch.Axes($ax).Delete() } catch {} }
        $pa = $ch.PlotArea
        foreach ($set in { $pa.Top = $pad }, { $pa.Left = $pad }, { $pa.Height = $h }, { $pa.Width = $w }, { $pa.InsideHeight = $h }, { $pa.InsideWidth = $w }) { try { & $set } catch {} }
        $ch.ChartGroups(1).GapWidth = 0; $ch.ChartGroups(1).Overlap = 100
        $s1 = $ch.SeriesCollection(1); $s1.Format.Fill.ForeColor.RGB = $color; $s1.Format.Line.Visible = 0
        $s2 = $ch.SeriesCollection(2); $s2.Format.Fill.ForeColor.RGB = $TRACK; $s2.Format.Fill.Transparency = 0; $s2.Format.Line.Visible = 0
    }
    # row 1: four small cards (rows 5-8); row 2: two large cards (rows 9-13)
    $dash.Range('5:8').RowHeight = 24; $dash.Range('9:13').RowHeight = 22
    for ($i = 0; $i -lt $kpis.Count; $i++) {
        $k = $kpis[$i]
        $cell = $pd.Cells(20 + $i, 2)
        $pd.Cells(20 + $i, 1).Value2 = $k[0]; $cell.Formula = $k[1]; $cell.NumberFormat = $k[2]
        $b = if ($i -lt 4) { Slot $i 4 'B5:B8' } else { Slot ($i - 4) 2 'B9:B13' }
        $bg = $dash.Shapes.AddShape(5, $b.L, $b.T, $b.W, $b.H)
        $bg.Adjustments.Item(1) = 0.12
        $bg.Fill.ForeColor.RGB = $WHITE; $bg.Line.ForeColor.RGB = $P2; $bg.Line.Weight = 0.75
        $bg.Shadow.Visible = -1; $bg.Shadow.Transparency = 0.7; $bg.Shadow.OffsetX = 2; $bg.Shadow.OffsetY = 2
        $bg.TextFrame2.TextRange.Text = ''
        Float $bg
        $row = 20 + $i
        switch ($k[0]) {
            'AVG PROGRESS' {
                $RD = 116; $gl = $b.L + 24; Add-Gauge $gl ($b.T + ($b.H - $RD) / 2 + 6) $RD $row $k[6]
                $pd.Cells($row, 7).Formula = '="across "&B20&" project"&IF(B20=1,"","s")'
                $tl = $gl + $RD + 4; $tw = $b.L + $b.W - $tl - 20
                Add-Text $tl $($b.T + 24) $tw 18 $k[0] 14 $true $INK | Out-Null
                Add-Text $tl $($b.T + 48) $tw 16 $k[4] 11 $false $SUB | Out-Null
                Add-Text $tl $($b.T + 68) $tw 16 '' 11 $true $SUB "=PivotData!`$G`$$row" | Out-Null
            }
            'COMPLETION RATE' {
                $pd.Cells($row, 7).Formula = '=B21&" of "&B20&" projects completed"'
                $tl = $b.L + 30; $tw = $b.W - 60
                Add-Text $tl $($b.T + 10) $tw 18 $k[0] 13 $true $INK | Out-Null
                Add-Text $tl $($b.T + 26) $tw 38 '' 30 $true $INK "=PivotData!`$B`$$row" | Out-Null
                Add-ProgressBar $tl ($b.T + 66) $tw 12 $row $k[6]
                Add-Text $tl $($b.T + $b.H - 20) $tw 14 '' 10 $false $SUB "=PivotData!`$G`$$row" | Out-Null
            }
            default {
                $ID = 40; Add-Icon ($b.L + 14) ($b.T + ($b.H - $ID) / 2) $ID $k[5] $k[6]
                $tl = $b.L + $ID + 20; $tw = $b.W - $ID - 28
                Add-Text $tl $($b.T + 6) $tw 22 $k[0] 16 $true $k[6] | Out-Null
                Add-Text $tl $($b.T + 29) $tw 36 '' 26 $true $INK "=PivotData!`$B`$$row" | Out-Null
                Add-Text $tl $($b.T + $b.H - 23) $tw 20 $k[4] 14 $false $SUB | Out-Null
            }
        }
    }

    # section headers are shapes (not cells) so they sit on top of the background picture
    function Section($row, $text) {
        $c = $dash.Range("B$row")
        $tb = $dash.Shapes.AddTextbox(1, $L0, $c.Top, 400, $c.Height)
        $tb.Fill.Visible = 0; $tb.Line.Visible = 0
        $tf = $tb.TextFrame2; $tf.MarginLeft = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0; $tf.VerticalAnchor = 4
        $tf.TextRange.Text = $text; $tf.TextRange.Font.Name = 'Segoe UI'; $tf.TextRange.Font.Size = 11
        $tf.TextRange.Font.Bold = 1; $tf.TextRange.Font.Fill.ForeColor.RGB = $INK
        Float $tb
        $y = $c.Top + $c.Height + 1
        $ln = $dash.Shapes.AddLine($L0, $y, $L0 + $W0, $y)
        $ln.Line.ForeColor.RGB = $A_VIOLET; $ln.Line.Weight = 2; Float $ln
    }
    function Box($addr) { $r = $dash.Range($addr); @{ L = $r.Left; T = $r.Top; W = $r.Width; H = $r.Height } }

    $all = @($ptKPI, $ptHealth, $ptCat, $ptLOB, $ptPOC)

    # Timeline slicer (Start Date)
    Section 15 'FILTERS'
    $tc = $wb.SlicerCaches.Add2($ptKPI, 'Start Date', 'Timeline_StartDate', 2)
    foreach ($p in $all) { if ($p.Name -ne 'ptKPI') { $tc.PivotTables.AddPivotTable($p) } }
    $b = Slot 0 1 'B16:B22'
    $tl = $tc.Slicers.Add($dash, $M, 'tlStartDate', 'Start Date', $b.T, $b.L, $b.W, $b.H)
    Float $tl.Shape

    # Field slicers
    $sl = @('Status', 'Health', 'LOB', 'Priority', 'Category', 'POC')
    for ($i = 0; $i -lt $sl.Count; $i++) {
        $sc = $wb.SlicerCaches.Add2($ptKPI, $sl[$i])
        foreach ($p in $all) { if ($p.Name -ne 'ptKPI') { $sc.PivotTables.AddPivotTable($p) } }
        $b = Slot $i $sl.Count 'B23:B34'
        $s = $sc.Slicers.Add($dash, $M, "sl$($sl[$i])", $sl[$i], $b.T, $b.L, $b.W, $b.H)
        $s.Style = 'SlicerStyleLight1'
        Float $s.Shape
    }

    # Charts
    Section 36 'ANALYTICS'
    function Add-PivotChart($pt, $type, $i, $rows, $title) {
        $b = Slot $i 2 $rows
        $shp = $dash.Shapes.AddChart2(-1, $type, $b.L, $b.T, $b.W, $b.H); Float $shp
        $ch = $shp.Chart
        $ch.SetSourceData($pt.TableRange1)
        $ch.HasTitle = $true; $ch.ChartTitle.Text = $title
        $ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
        $ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = $true
        $ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $INK
        try { $ch.ShowAllFieldButtons = $false } catch {}
        $ch
    }
    $c1 = Add-PivotChart $ptHealth -4120 0 'B37:B53' 'Project Health'
    $c1.SeriesCollection(1).HasDataLabels = $true
    $c1.SeriesCollection(1).DataLabels().Font.Size = 11; $c1.SeriesCollection(1).DataLabels().Font.Bold = $true
    $c1.HasLegend = $true; $c1.Legend.Position = -4107
    $c2 = Add-PivotChart $ptLOB 52 1 'B37:B53' 'Projects by LOB & Status'
    $c2.HasLegend = $true; $c2.Legend.Position = -4107
    $c3 = Add-PivotChart $ptPOC 58 0 'B54:B70' 'Workload by POC & Status'
    $c3.HasLegend = $true; $c3.Legend.Position = -4107
    $c3.Axes(2).MajorUnit = 1
    $c4 = Add-PivotChart $ptCat 57 1 'B54:B70' 'Average Progress by Category'
    $c4.HasLegend = $false
    $c4.Axes(2).MinimumScale = 0; $c4.Axes(2).MaximumScale = 1; $c4.Axes(2).TickLabels.NumberFormat = '0%'
    $c4.SeriesCollection(1).HasDataLabels = $true
    $dl = $c4.SeriesCollection(1).DataLabels(); $dl.NumberFormat = '0%'; $dl.Position = 3
    $dl.Font.Color = $WHITE; $dl.Font.Bold = $true
    $c1.SeriesCollection(1).DataLabels().Font.Color = $WHITE

    # ============================================================ TIMELINE (GANTT)
    $g = $gantt
    $g.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
    $g.Cells.Font.Name = 'Segoe UI'; $g.Cells.Font.Size = 9
    $g.Range('B1').Value2 = 'Project Timeline (Gantt)'
    $g.Range('B1').Font.Size = 16; $g.Range('B1').Font.Bold = $true; $g.Range('B1').Font.Color = $INK
    $g.Range('B3').Value2 = 'Search:'; $g.Range('E3').Value2 = 'Status:'; $g.Range('B4').Value2 = 'Timeline start:'
    $g.Range('B3:B4').Font.Bold = $true; $g.Range('E3').Font.Bold = $true
    $g.Range('B3').HorizontalAlignment = -4152; $g.Range('E3').HorizontalAlignment = -4152; $g.Range('B4').HorizontalAlignment = -4152
    foreach ($a in 'C3', 'F3', 'C4') { $g.Range($a).Interior.Color = $P4; $g.Range($a).BorderAround(1, 2) | Out-Null }
    $g.Range('F3').Value2 = 'All'
    $g.Range('F3').Validation.Add(3, 1, 1, 'All,Completed,Ongoing,Not Started,On Hold') | Out-Null
    $g.Range('C4').Formula = '=MIN(tblProjects[Start Date])-WEEKDAY(MIN(tblProjects[Start Date]),3)'
    $g.Range('C4').NumberFormat = 'm/d/yyyy'
    $g.Range('D3').Value2 = '<- type project, POC, developer or LOB'
    $g.Range('D4').Value2 = '<- override with any date to scroll the chart'
    $g.Range('D3:D4').Font.Italic = $true; $g.Range('D3:D4').Font.Color = $GREY

    # legend
    $G_DONE = $A_VIOLET; $G_LEFT = HEX '#7DD3FC'; $G_SLIP = HEX '#FDA4AF'; $G_TODAY = HEX '#FDE68A'
    $leg = @(@('I5', $G_DONE, 'Done'), @('M5', $G_LEFT, 'Remaining'), @('S5', $G_SLIP, 'Overdue slip'), @('Z5', $G_TODAY, 'Today'))
    foreach ($l in $leg) { $g.Range($l[0]).Interior.Color = $l[1]; $g.Range($l[0]).Offset(0, 1).Value2 = $l[2] }

    $hdrs = 'Project', 'POC', 'Status', 'Start', 'Due', 'Progress', 'Show'
    for ($i = 0; $i -lt $hdrs.Count; $i++) { $g.Cells(7, 2 + $i).Value2 = $hdrs[$i] }
    $nRows = 100; $r0 = 8; $rN = $r0 + $nRows - 1
    $nDays = 91; $d0 = 9; $dN = $d0 + $nDays - 1
    $idx = 'ROWS($B$8:$B8)'
    $g.Range("B$($r0):B$rN").Formula = "=IFERROR(INDEX(tblProjects[Project Name],$idx),`"`")"
    $g.Range("C$($r0):C$rN").Formula = "=IFERROR(INDEX(tblProjects[POC],$idx),`"`")"
    $g.Range("D$($r0):D$rN").Formula = "=IFERROR(INDEX(tblProjects[Status],$idx),`"`")"
    $g.Range("E$($r0):E$rN").Formula = "=IFERROR(INDEX(tblProjects[Start Date],$idx),`"`")"
    $g.Range("F$($r0):F$rN").Formula = "=IFERROR(INDEX(tblProjects[Due Date],$idx),`"`")"
    $g.Range("G$($r0):G$rN").Formula = "=IFERROR(INDEX(tblProjects[Progress],$idx),`"`")"
    $g.Range("H$($r0):H$rN").Formula = "=IFERROR(AND(`$B8<>`"`",OR(`$C`$3=`"`",ISNUMBER(SEARCH(`$C`$3,`$B8&`"|`"&`$C8&`"|`"&INDEX(tblProjects[Developer],$idx)&`"|`"&INDEX(tblProjects[LOB],$idx)))),OR(`$F`$3=`"All`",`$F`$3=`$D8)),FALSE)"
    $g.Range("E$($r0):F$rN").NumberFormat = 'm/d/yy'
    $g.Range("G$($r0):G$rN").NumberFormat = '0%'
    $g.Columns.Item('H').Hidden = $true
    $g.Columns.Item('A').ColumnWidth = 2; $g.Columns.Item('B').ColumnWidth = 34; $g.Columns.Item('C').ColumnWidth = 18
    $g.Columns.Item('D').ColumnWidth = 11; $g.Range('E:F').ColumnWidth = 8; $g.Columns.Item('G').ColumnWidth = 8

    $g.Cells(7, $d0).Formula = '=$C$4'
    $g.Range($g.Cells(7, $d0 + 1), $g.Cells(7, $dN)).Formula = "=$($g.Cells(7, $d0).Address($false, $false))+1"
    $g.Range($g.Cells(7, $d0), $g.Cells(7, $dN)).NumberFormat = 'd'
    for ($c = $d0; $c -le $dN; $c += 7) {
        $wk = $g.Range($g.Cells(6, $c), $g.Cells(6, [Math]::Min($c + 6, $dN)))
        $wk.Merge(); $wk.Formula = "=TEXT($($g.Cells(7, $c).Address($false, $false)),`"mmm d`")"
        $wk.Borders.Item(7).Color = (RGB 191 191 191)
    }
    $g.Range($g.Cells(6, $d0), $g.Cells(6, $dN)).Font.Bold = $true
    $g.Range($g.Cells(6, $d0), $g.Cells(6, $dN)).HorizontalAlignment = -4131
    $g.Range($g.Cells(7, $d0), $g.Cells(7, $dN)).HorizontalAlignment = -4108
    $g.Range($g.Cells(1, $d0), $g.Cells(1, $dN)).ColumnWidth = 2.6

    $hdrRow = $g.Range($g.Cells(7, 2), $g.Cells(7, $dN))
    $hdrRow.Interior.Color = $P2; $hdrRow.Font.Color = $INK; $hdrRow.Font.Bold = $true
    $body = $g.Range($g.Cells($r0, 2), $g.Cells($rN, $dN))
    $body.Borders.Item(12).Color = (RGB 217 217 217); $body.Borders.Item(12).LineStyle = 1
    $body.Borders.Item(9).Color = (RGB 217 217 217); $body.Borders.Item(9).LineStyle = 1

    # conditional formats for the Gantt grid (relative to top-left cell)
    $g.Cells($r0, $d0).Select() | Out-Null
    $grid = $g.Range($g.Cells($r0, $d0), $g.Cells($rN, $dN))
    $D = $g.Cells(7, $d0).Address($true, $false)   # e.g. I$7
    $rules = @(
        @("=AND(`$H8,$D>=`$E8,$D<=`$F8,$D<`$E8+(`$F8-`$E8+1)*`$G8)", $G_DONE),
        @("=AND(`$H8,$D>=`$E8,$D<=`$F8)", $G_LEFT),
        @("=AND(`$H8,`$D8<>`"Completed`",$D>`$F8,$D<=TODAY())", $G_SLIP),
        @("=$D=TODAY()", $G_TODAY),
        @("=WEEKDAY($D,2)>5", (RGB 242 242 242))
    )
    foreach ($rule in $rules) {
        $fc = $grid.FormatConditions.Add(2, $M, $rule[0])
        $fc.Interior.Color = $rule[1]; $fc.SetLastPriority() | Out-Null
    }
    $g.Range("B$($r0)").Select() | Out-Null
    $info = $g.Range("B$($r0):G$rN")
    $fc = $info.FormatConditions.Add(2, $M, '=AND($B8<>"",NOT($H8))'); $fc.Font.Color = (RGB 191 191 191)
    $gdb = $g.Range("G$($r0):G$rN").FormatConditions.AddDatabar()
    $gdb.MinPoint.Modify(0, 0); $gdb.MaxPoint.Modify(0, 1); $gdb.BarColor.Color = $A_EMERALD
    Add-StatusCF $g.Range("D$($r0):D$rN") 'Completed'
    Add-StatusCF $g.Range("D$($r0):D$rN") 'Ongoing'
    $g.Columns.Item(1).Insert() | Out-Null
    $g.Range('1:1').RowHeight = 30; $g.Range('2:5').RowHeight = 28; $g.Range('6:7').RowHeight = 20
    Add-Nav $g 'Timeline'
    $g.Activate(); $g.Range('J8').Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true

    # ============================================================ DROPDOWN LISTS (hidden)
    $defaults = [ordered]@{
        'POC'               = @()
        'LOB'               = @()
        'Status'            = @('Not Started', 'Ongoing', 'Completed', 'On Hold')
        'Priority'          = @('High', 'Medium', 'Low')
        'Category'          = @()
        'Deployment Status' = @('Not Yet Started', 'Submitted', 'In UAT', 'Deployed')
    }
    $lc = 1; $listVals = @{}
    foreach ($name in $defaults.Keys) {
        $ci = [array]::IndexOf($cols, $name)
        $vals = @($defaults[$name]) + @($rows | ForEach-Object { $_[$ci] }) | Where-Object { $_ } | Select-Object -Unique
        if ($defaults[$name].Count -eq 0) { $vals = $vals | Sort-Object }
        $listVals[$name] = @($vals)
        $lists.Cells(1, $lc).Value2 = $name
        $r = 2; foreach ($v in $vals) { $lists.Cells($r, $lc).Value2 = $v; $r++ }
        $colL = $lists.Cells(1, $lc).Address($false, $false) -replace '\d', ''
        $wb.Names.Add("lst_$($name -replace ' ', '')", "=OFFSET(Lists!`$$colL`$2,0,0,MAX(1,COUNTA(Lists!`$$($colL):`$$colL)-1),1)") | Out-Null
        $lc++
    }
    $lists.Range('1:1').Font.Bold = $true
    $lists.Visible = 0

    # ============================================================ PROJECTS TABLE DROPDOWNS (for manual adds)
    $listOf = @{ 'POC' = 'lst_POC'; 'LOB' = 'lst_LOB'; 'Status' = 'lst_Status'; 'Priority' = 'lst_Priority'; 'Category' = 'lst_Category'; 'Deployment Status' = 'lst_DeploymentStatus' }
    foreach ($name in $listOf.Keys) {
        $v = $lo.ListColumns.Item($name).DataBodyRange.Validation
        $v.Delete(); $v.Add(3, 1, 1, "=$($listOf[$name])") | Out-Null; $v.ShowError = $false
    }
    foreach ($name in 'Start Date', 'Due Date') {
        $v = $lo.ListColumns.Item($name).DataBodyRange.Validation
        $v.Delete(); $v.Add(4, 1, 1, '=DATE(2000,1,1)', '=DATE(2100,12,31)') | Out-Null
        $v.ErrorMessage = 'Please enter a valid date (m/d/yyyy).'
    }
    $v = $lo.ListColumns.Item('Progress').DataBodyRange.Validation
    $v.Delete(); $v.Add(2, 1, 1, '0', '1') | Out-Null; $v.ErrorMessage = 'Progress must be between 0% and 100%.'

    # ============================================================ NEWSLETTER (fully editable)
    $n = $news
    $n.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
    $n.Cells.Font.Name = 'Segoe UI'; $n.Cells.Font.Size = 10; $n.Cells.Font.Color = $INK
    $n.Columns.Item('A').ColumnWidth = 2; $n.Columns.Item('B').ColumnWidth = 5; $n.Columns.Item('C').ColumnWidth = 18
    $n.Columns.Item('D').ColumnWidth = 40; $n.Columns.Item('E').ColumnWidth = 85; $n.Columns.Item('F').ColumnWidth = 30; $n.Columns.Item('G').ColumnWidth = 2
    $n.Range('A1:G3').Interior.Color = $P1
    $n.Range('A4:G4').Interior.Color = $P3; $n.Range('4:4').RowHeight = 6
    $n.Range('B1').Value2 = 'IGNITE Project Newsletter'
    $n.Range('B1').Font.Size = 22; $n.Range('B1').Font.Bold = $true; $n.Range('1:1').RowHeight = 38
    $n.Range('B2').Value2 = 'Issue #1  |  October 2026  |  Payment Integrity - Dispute Intake - Non-COB'
    $n.Range('B2').Font.Size = 11
    $n.Range('B3').Value2 = 'Prepared by: [Your name]'
    $n.Range('B3').Font.Size = 10; $n.Range('B3').Font.Italic = $true

    $n.Range('B6:F6').Merge()
    $n.Range('B6').Value2 = "Welcome to our project newsletter! Here's a quick look at what we delivered, what's in progress, and what needs our attention. Every line on this page can be edited - make it yours."
    $n.Range('B6').WrapText = $true; $n.Range('6:6').RowHeight = 34; $n.Range('B6').Font.Size = 11; $n.Range('B6').VerticalAlignment = -4108

    $n.Range('B8:F8').Merge()
    $n.Range('B8').Formula = '="At a glance:   "&ROWS(tblProjects)&" projects   |   "&COUNTIFS(tblProjects[Status],"Completed")&" completed   |   "&COUNTIFS(tblProjects[Status],"Ongoing")&" ongoing   |   "&SUM(tblProjects[Is Overdue])&" overdue   |   "&TEXT(AVERAGE(tblProjects[Progress]),"0%")&" average progress"'
    $n.Range('B8').Interior.Color = $P4; $n.Range('B8').Font.Bold = $true; $n.Range('B8').Font.Size = 11
    $n.Range('8:8').RowHeight = 28; $n.Range('B8').VerticalAlignment = -4108; $n.Range('B8').IndentLevel = 1

    $n.Range('B10:F10').Merge()
    $n.Range('B10').Value2 = 'Tip: everything here is editable. Add an item by typing in the first empty row below the table. Remove an item: right-click the row > Delete > Table Rows. Pick a Section from the dropdown or type your own.'
    $n.Range('B10').Font.Italic = $true; $n.Range('B10').Font.Color = $SUB; $n.Range('B10').WrapText = $true; $n.Range('10:10').RowHeight = 28

    $items = @(
        @('Highlights', '4 projects delivered', 'NTE and NOD tracker (VANGUARD - Power Apps), Attendance (with PTO request) tracker (Tempus - Power Apps), OPIDs Review and Genpro are now complete. Thank you to everyone who made these possible!', 'John Ryan Ariola, Charlon Calugay, Justin Umali, Pope Tuazon'),
        @('Spotlight', 'Tempus: attendance and PTO requests in one place', 'The new Attendance tracker lets the team log attendance and file PTO requests in a single Power App, replacing manual tracking. Try it out and share your feedback.', 'Charlon Calugay'),
        @('In Progress', 'Employee Dashboard - 50% complete', 'A one-stop view of team metrics for managers. Target completion: Sep 4, 2026.', 'John Ryan Ariola'),
        @('In Progress', 'New Ansir Apps and Ansir Calculator for Non-COB - 75%', 'Two new tools to speed up Non-COB work are in final development. Target completion: Sep 1, 2026.', 'Shine Abdulhalik, Junie Ann Martinez'),
        @('In Progress', 'Root Cause Analysis tool - 75%', 'A process-improvement tool to track and analyze root causes. Target completion: Sep 11, 2026.', 'John Ryan Ariola'),
        @('In Progress', 'Payment Integrity - Pends Review (APEX)', 'Ageing-management app for pends review. Development is done and the project is being wrapped up.', 'John Ryan Ariola'),
        @('Needs Attention', '5 projects are past their due date', 'Employee Dashboard, New Ansir Apps for NonCOB, Ansir Calculator for NonCOB, RootCauseAnalysis and Payment Integrity - Pends Review are still Ongoing past their due dates. Owners: please update the status or set a new target date in the Projects tab.', 'Project POCs'),
        @('Coming Up', 'Next milestones', 'Wrap up the remaining Non-COB tools, launch the Employee Dashboard, and line up the next automation ideas. Have an idea? Share it with the team!', 'Team'),
        @('Team Shout-out', 'Thank you, developers!', 'Big thanks to Shine, Junie Ann, Pope, Justin, Charmaine, Charlon and Melody for building and supporting this cycle''s apps.', 'Everyone'),
        @('Reminders', 'Keep the Project Hub up to date', 'Add new projects in the Projects tab (first empty row below the table), then click Data > Refresh All so the Dashboard shows the latest numbers.', 'All POCs'),
        @('Contact', 'Questions or ideas?', 'Reach out to the newsletter editor or any of the POCs listed above.', '[Editor name]')
    )
    $nh = 12
    $hdrN = '#', 'Section', 'Headline', 'Details', 'Owner / Contact'
    for ($i = 0; $i -lt 5; $i++) { $n.Cells($nh, 2 + $i).Value2 = $hdrN[$i] }
    for ($r = 0; $r -lt $items.Count; $r++) { for ($c2 = 0; $c2 -lt 4; $c2++) { $n.Cells($nh + 1 + $r, 3 + $c2).Value2 = $items[$r][$c2] } }
    $nlo = $n.ListObjects.Add(1, $n.Range($n.Cells($nh, 2), $n.Cells($nh + $items.Count, 6)), $M, 1)
    $nlo.Name = 'tblNewsletter'; $nlo.TableStyle = 'TableStyleMedium2'; $nlo.ShowAutoFilterDropDown = $false
    $nlo.HeaderRowRange.Font.Color = $WHITE; $nlo.HeaderRowRange.RowHeight = 24
    $nlo.ListColumns.Item('#').DataBodyRange.Formula = '=ROW()-ROW(tblNewsletter[#Headers])'
    $nlo.ListColumns.Item('#').Range.HorizontalAlignment = -4108
    $nlo.ListColumns.Item('Headline').DataBodyRange.Font.Bold = $true
    $nlo.DataBodyRange.WrapText = $true; $nlo.DataBodyRange.VerticalAlignment = -4160
    $sec = $nlo.ListColumns.Item('Section').DataBodyRange
    $sv = $sec.Validation; $sv.Delete()
    $sv.Add(3, 1, 1, 'Highlights,Spotlight,In Progress,Needs Attention,Coming Up,Team Shout-out,Reminders,Contact') | Out-Null; $sv.ShowError = $false
    $secColors = [ordered]@{
        'Highlights' = $STC['Completed']; 'In Progress' = $STC['Ongoing']; 'Needs Attention' = $STC['Overdue']; 'Coming Up' = $STC['On Track']
        'Spotlight' = @((HEX '#EDE9FE'), (HEX '#5B21B6')); 'Team Shout-out' = @((HEX '#CCFBF1'), (HEX '#115E59'))
        'Reminders' = @($P2, $INK); 'Contact' = @((HEX '#F1F2F4'), $SUB)
    }
    foreach ($s in $secColors.Keys) { Add-TextCF $sec $s $secColors[$s][0] $secColors[$s][1] }
    $nlo.DataBodyRange.Rows.AutoFit() | Out-Null
    try {
        $ps = $n.PageSetup; $ps.Orientation = 2; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false
    } catch {}
    $n.Columns.Item(1).Insert() | Out-Null
    Add-Nav $n 'Newsletter'
    $n.Range('C2').Select() | Out-Null

    # ============================================================ finish
    # dashboard background picture behind all shapes (content area: column B..Q, row 4..72)
    $bgL = $dash.Range('B4').Left; $bgT = $dash.Range('B4').Top
    $bgW = $dash.Range('B4:Q4').Width; $bgH = $dash.Range('A4:A72').Height
    $bgFile = Join-Path $env:TEMP "dash_bg_$PID.jpg"
    New-BackgroundImage $bgFile ([int]($bgW * 1.5)) ([int]($bgH * 1.5))
    $bgS = $dash.Shapes.AddShape(1, $bgL, $bgT, $bgW, $bgH)
    $bgS.Name = 'bg_canvas'; $bgS.Line.Visible = 0
    $bgS.Fill.UserPicture($bgFile)
    $bgS.ZOrder(1); $bgS.Placement = 3
    Remove-Item $bgFile -ErrorAction SilentlyContinue

    $pd.Range('A1').Value2 = 'Helper pivots for Dashboard - do not edit'
    $pd.Visible = 0
    $proj.Activate(); $proj.Range('B4').Select() | Out-Null
    $dash.Activate(); $dash.Range('B4').Select() | Out-Null
    $xl.ActiveWindow.DisplayWorkbookTabs = $false
    $xl.Calculation = -4105
    $wb.RefreshAll()
    if (Test-Path $OutFile) { Remove-Item $OutFile }
    $wb.SaveAs($OutFile, 51)
    "Saved: $OutFile"
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.ScreenUpdating = $true
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
