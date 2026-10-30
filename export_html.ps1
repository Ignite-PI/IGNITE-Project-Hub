param(
    [string]$Workbook = "$PSScriptRoot\Project_Consolidation_Dashboard.xlsx",
    [string]$Template = "$PSScriptRoot\dashboard_template.html",
    [string]$OutFile = "$PSScriptRoot\Project_Hub.html",
    [switch]$RemoveLoginPassword   # use this if the login password is forgotten
)
# Reads the Projects table and the Newsletter from the Excel workbook and writes a single,
# self-contained interactive HTML file (no internet or libraries needed).
$ErrorActionPreference = 'Stop'
function IsoDate($v) {
    if ($v -is [double]) { return [datetime]::FromOADate($v).ToString('yyyy-MM-dd') }
    if ($v -is [datetime]) { return $v.ToString('yyyy-MM-dd') }
    return ''
}
function ColValues($lo, $name) {
    $x = $lo.ListColumns.Item($name).DataBodyRange.Value2
    if ($x -is [array]) { return , $x } else { $a = New-Object 'object[,]' 1, 1; $a[0, 0] = $x; return , $a }
}
function At($arr, $i) { $arr.GetValue($i + $arr.GetLowerBound(0), $arr.GetLowerBound(1)) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Workbook, 0, $true)
    $lo = $null
    for ($w = 0; $w -lt 60 -and -not $lo; $w++) {
        try { $lo = $wb.Worksheets.Item('Projects').ListObjects.Item('tblProjects') } catch { Start-Sleep 2 }
    }
    $map = [ordered]@{ name = 'Project Name'; poc = 'POC'; dev = 'Developer'; lob = 'LOB'; status = 'Status'; priority = 'Priority'
        category = 'Category'; start = 'Start Date'; due = 'Due Date'; deploy = 'Deployment Status'; progress = 'Progress'; link = 'Link' }
    $vals = @{}; foreach ($k in $map.Keys) { $vals[$k] = ColValues $lo $map[$k] }
    $projects = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $lo.ListRows.Count; $i++) {
        $p = [ordered]@{}
        foreach ($k in $map.Keys) {
            $v = At $vals[$k] $i
            $p[$k] = switch ($k) { { $_ -in 'start', 'due' } { IsoDate $v } 'progress' { if ($v -is [double]) { $v } else { 0 } } default { if ($null -eq $v) { '' } else { "$v".Trim() } } }
        }
        if ($p.name) { $projects.Add([pscustomobject]$p) }
    }

    $ns = $wb.Worksheets.Item('Newsletter')
    $nlo = $ns.ListObjects.Item('tblNewsletter')
    $nmap = [ordered]@{ section = 'Section'; headline = 'Headline'; details = 'Details'; owner = 'Owner / Contact' }
    $nv = @{}; foreach ($k in $nmap.Keys) { $nv[$k] = ColValues $nlo $nmap[$k] }
    $items = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $nlo.ListRows.Count; $i++) {
        $it = [ordered]@{}; foreach ($k in $nmap.Keys) { $v = At $nv[$k] $i; $it[$k] = if ($null -eq $v) { '' } else { "$v" } }
        if ($it.section -or $it.headline -or $it.details) { $items.Add([pscustomobject]$it) }
    }
    $newsletter = [ordered]@{
        title = $ns.Range('C1').Text; issue = $ns.Range('C2').Text; preparedBy = $ns.Range('C3').Text; intro = $ns.Range('C6').Text
        items = $items.ToArray()
    }
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

$data = [ordered]@{
    generated  = (Get-Date).ToString('MMMM d, yyyy h:mm tt')
    projects   = $projects.ToArray()
    newsletter = $newsletter
}
$json = ($data | ConvertTo-Json -Depth 6 -Compress).Replace('</', '<\/')
$html = [IO.File]::ReadAllText($Template).Replace('/*__DATA__*/null', $json)
# keep the edit password that is built into the current HTML (set with "Save protected copy")
if (Test-Path $OutFile) {
    $m = [regex]::Match([IO.File]::ReadAllText($OutFile), "const EDIT_PW_HASH = /\*__PW__\*/'([0-9a-f]*)';")
    if ($m.Success -and $m.Groups[1].Value) { $html = $html.Replace("/*__PW__*/''", "/*__PW__*/'$($m.Groups[1].Value)'") }
    # keep the login password too (unless -RemoveLoginPassword)
    $v = [regex]::Match([IO.File]::ReadAllText($OutFile), "const VIEW_PW_HASH = /\*__VPW__\*/'([0-9a-f]*)';")
    if (-not $RemoveLoginPassword -and $v.Success -and $v.Groups[1].Value) { $html = $html.Replace("/*__VPW__*/''", "/*__VPW__*/'$($v.Groups[1].Value)'") }
    if ($RemoveLoginPassword) { Write-Host 'Login password removed from Project_Hub.html' }
}
[IO.File]::WriteAllText($OutFile, $html, (New-Object System.Text.UTF8Encoding $false))
"Exported $($projects.Count) projects and $($items.Count) newsletter items -> $OutFile"
