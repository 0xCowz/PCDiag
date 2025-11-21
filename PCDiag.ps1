$outputPath = if (Test-Path "$env:USERPROFILE\OneDrive - Humens Bidco SAS\Documents\Rapports Diag Hebdomadaire") { "$env:USERPROFILE\OneDrive - Humens Bidco SAS\Documents\Rapports Diag Hebdomadaire" } else { "$env:USERPROFILE\Documents\Rapports Diag Hebdomadaire" }

if (-not (Test-Path $outputPath)) {
    mkdir $outputPath -Force > $null
}

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

Write-Host "operation 1 en cours, veuillez patienter.."

$report = @()
$last24h = (Get-Date).AddHours(-24)

$errors = Get-WinEvent -FilterHashtable @{LogName='System'; Level=2; StartTime=$last24h} -MaxEvents 10 -ErrorAction SilentlyContinue
if ($errors -and $errors.Count -gt 0) {
  foreach ($e in $errors) {
    $report += [pscustomobject]@{ Section="Event Viewer"; Etat="ERREUR"; Details="$( $e.TimeCreated ) - $( $e.ProviderName )" }
  }
} else {
  $report += [pscustomobject]@{ Section="Event Viewer"; Etat="OK"; Details="Aucune erreur (24h)" }
}

$stopped = Get-Service | Where-Object { $_.Status -eq "Stopped" -and $_.StartType -eq "Automatic" }
if ($stopped) {
  foreach ($svc in $stopped) {
    $report += [pscustomobject]@{ Section="Services"; Etat="ERREUR"; Details="$( $svc.DisplayName ) ($( $svc.Name ))" }
  }
} else {
  $report += [pscustomobject]@{ Section="Services"; Etat="OK"; Details="Tous les services sont actifs" }
}

$accounts = Get-ItemProperty -Path "HKCU:\Software\Microsoft\OneDrive\Accounts\*" -ErrorAction SilentlyContinue
if ($accounts) {
  foreach ($acc in $accounts) {
    $folder = $acc.UserFolder
    $name   = if ($acc.DisplayName) { $acc.DisplayName } else { "OneDrive" }
    if (Test-Path $folder) {
      $lastFile = Get-ChildItem -Path $folder -Recurse -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
      if ($lastFile) {
        $report += [pscustomobject]@{ Section="OneDrive ($name)"; Etat="OK"; Details="Dernier sync: $( $lastFile.LastWriteTime ) - $( $lastFile.Name )" }
      } else {
        $report += [pscustomobject]@{ Section="OneDrive ($name)"; Etat="ALERTE"; Details="Aucun fichier trouve" }
      }
    } else {
      $report += [pscustomobject]@{ Section="OneDrive ($name)"; Etat="ERREUR"; Details="Dossier introuvable" }
    }
  }
} else {
  $report += [pscustomobject]@{ Section="OneDrive"; Etat="ERREUR"; Details="Aucun compte detecte" }
}

$drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -ge 0 -and $_.Free -ge 0 }
foreach ($d in $drives) {
  $totalBytes = $d.Used + $d.Free
  if ($totalBytes -eq 0) { continue }
  $totalGB = [math]::Round($totalBytes/1GB,2)
  $freeGB  = [math]::Round($d.Free/1GB,2)
  $percentFree = [math]::Round(($d.Free / $totalBytes) * 100,2)
  if ($percentFree -lt 5) { $etat = "CRITIQUE" }
  elseif ($percentFree -lt 10) { $etat = "ALERTE" }
  else { $etat = "OK" }
  $report += [pscustomobject]@{ Section="Disque $( $d.Name ):"; Etat=$etat; Details="Total: $totalGB Go - Libre: $freeGB Go ($percentFree %)" }
  $targetPath = "$( $d.Name ):\"
  if (Test-Path $targetPath) {
    $folderSizes = @()
    foreach ($folder in Get-ChildItem -Path $targetPath -Directory -Force) {
      try {
        $size = (Get-ChildItem -Path $folder.FullName -File -ErrorAction Stop | Measure-Object -Property Length -Sum).Sum
        foreach ($sub in Get-ChildItem -Path $folder.FullName -Directory -Force -ErrorAction SilentlyContinue) {
          $size += (Get-ChildItem -Path $sub.FullName -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        }
        $folderSizes += [pscustomobject]@{ Dossier=$folder.FullName; TailleGB=[math]::Round(($size/1GB),2) }
      } catch {}
    }
    if ($folderSizes.Count -gt 0) {
      foreach ($f in ($folderSizes | Sort-Object -Property TailleGB -Descending | Select-Object -First 5)) {
        $report += [pscustomobject]@{ Section="Dossiers lourds $( $d.Name ):"; Etat="INFO"; Details="$( $f.Dossier ) - $( $f.TailleGB ) Go" }
      }
    }
  }
}

if (Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue) {
  foreach ($disk in Get-PhysicalDisk) {
    $report += [pscustomobject]@{ Section="Etat disques"; Etat=$disk.HealthStatus; Details="$( $disk.FriendlyName ) - $( $disk.OperationalStatus )" }
  }
}

try {
  $updateSession = New-Object -ComObject Microsoft.Update.Session
  $updateSearcher = $updateSession.CreateUpdateSearcher()
  $searchResult = $updateSearcher.Search("IsInstalled=0")
  if ($searchResult.Updates.Count -gt 0) {
    foreach ($update in $searchResult.Updates | Select-Object -First 5) {
      $report += [pscustomobject]@{ Section="Windows Update"; Etat="ALERTE"; Details=$update.Title }
    }
  } else {
    $report += [pscustomobject]@{ Section="Windows Update"; Etat="OK"; Details="Systeme a jour" }
  }
} catch {
  $report += [pscustomobject]@{ Section="Windows Update"; Etat="INFO"; Details="Impossible d interroger Windows Update (droits/reseau ?)" }
}

Add-Type -AssemblyName System.Web
function HtmlEncode([string]$s) { [System.Web.HttpUtility]::HtmlEncode($s) }
function StatusClass([string]$etat) {
  switch ($etat.ToString().ToUpper()) {
    'OK'       { 'ok' }
    'ALERTE'   { 'warn' }
    'ERREUR'   { 'error' }
    'CRITIQUE' { 'critical' }
    'INFO'     { 'info' }
    default    { 'info' }
  }
}
function StatusIcon([string]$etat) {
  switch ($etat.ToString().ToUpper()) {
    'OK'       { '&#10003;' }
    'ALERTE'   { '&#9888;' }
    'ERREUR'   { '&#10006;' }
    'CRITIQUE' { '&#9940;' }
    'INFO'     { '&#8505;' }
    default    { '&#8226;' }
  }
}

$culture   = [System.Globalization.CultureInfo]::GetCultureInfo('fr-FR')
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$prettyTs  = (Get-Date).ToString('dd MMMM yyyy HH:mm:ss', $culture)
$machine   = $env:COMPUTERNAME
$userId    = "$env:USERDOMAIN\$env:USERNAME"
$htmlPath  = "$outputPath\Rapport $timestamp.html"

$byEtat = @{}
$report | Group-Object Etat | ForEach-Object { $byEtat[$_.Name] = $_.Count }
$etatOrder = @('CRITIQUE','ERREUR','ALERTE','OK','INFO')
$summaryBadges = foreach ($e in $etatOrder) {
  $count = if ($byEtat.ContainsKey($e)) { $byEtat[$e] } else { 0 }
  $cls   = StatusClass $e
  "<div class='chip $cls'><span class='icon'>$(StatusIcon $e)</span>$e<span class='count'>$count</span></div>"
}
$summaryHtml = ($summaryBadges -join "`n")

$sectionBlocks = foreach ($grp in ($report | Group-Object Section | Sort-Object Name)) {
  $plural = if ($grp.Count -gt 1) { 's' } else { '' }
@"
<section class='section'>
  <h2><span class='caret'>&#9656;</span> $(HtmlEncode($grp.Name)) <span class='meta'>($( $grp.Count ) element$plural)</span></h2>
  <div class='rows'>
"@
  foreach ($item in $grp.Group) {
    $cls   = StatusClass $item.Etat
    $icon  = StatusIcon  $item.Etat
    $etat  = HtmlEncode($item.Etat)
    $det   = HtmlEncode($item.Details)
@"
    <div class='row $cls'>
      <div class='badge'><span class='icon'>$icon</span>$etat</div>
      <div class='detail'>$det</div>
    </div>
"@
  }
@"
  </div>
</section>
"@
}
$sectionsHtml = ($sectionBlocks -join "`n")

$styles = @'
:root{--bg:#f6f7fb;--card:#fff;--text:#111827;--muted:#6b7280;--border:#e5e7eb;--accent:#6366f1;--ok:#10b981;--warn:#f59e0b;--err:#ef4444;--crit:#b91c1c;--info:#3b82f6;}
*{box-sizing:border-box}html,body{margin:0;padding:0;height:100%}
body{background:var(--bg);color:var(--text);font-family:Segoe UI,Roboto,Arial,sans-serif;line-height:1.45;position:relative;min-height:100%}
body::before{content:"";position:fixed;top:0;left:0;right:0;bottom:0;background:url("https://www.humens.com/wp-content/uploads/2023/03/novawood-2023-scaled.jpg") no-repeat center center/cover;z-index:-1;filter:blur(6px) brightness(0.8);transform:scale(1.05)}
.header{background:transparent;color:#fff;padding:20px 0;text-align:center}
.header h1{margin:0 0 6px 0;font-size:26px;font-weight:700}
.header .meta{opacity:.95;font-size:14px}
.container{max-width:1100px;margin:-20px auto 40px auto;padding:0 16px}
.card{background:rgba(255,255,255,0.85);backdrop-filter:blur(4px);border:1px solid var(--border);border-radius:16px;box-shadow:0 6px 18px rgba(0,0,0,.06);padding:18px 20px;margin-bottom:18px}
.tools{display:flex;gap:12px;flex-wrap:wrap;align-items:center}
.search{flex:1;min-width:200px}
.search input{width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:10px;font-size:14px}
.summary{display:flex;gap:10px;flex-wrap:wrap}
.chip{display:inline-flex;align-items<center;gap:8px;border:1px solid var(--border);background:rgba(255,255,255,0.8);padding:8px 12px;border-radius:999px;font-size:13px}
.chip .count{margin-left:6px;background:rgba(0,0,0,.06);padding:2px 8px;border-radius:999px;font-weight:700}
.chip .icon{font-weight:700}
.chip.ok{color:var(--ok)}.chip.warn{color:var(--warn)}.chip.error{color:var(--err)}.chip.critical{color:var(--crit)}.chip.info{color:var(--info)}
.section{border:1px solid var(--border);border-radius:14px;background:rgba(255,255,255,0.85);backdrop-filter:blur(4px);margin-bottom:14px;overflow:hidden}
.section h2{margin:0;padding:14px 16px;background:rgba(250,250,250,0.85);border-bottom:1px solid var(--border);font-size:16px;display:flex;align-items:center;gap:8px;cursor:pointer}
.section .meta{color:var(--muted);font-weight:500}
.section .rows{padding:6px 12px}
.section.collapsed .rows{display:none}
.section .caret{transition:transform .2s ease}
.section.collapsed .caret{transform:rotate(0deg)}
.section:not(.collapsed) .caret{transform:rotate(90deg)}
.row{display:grid;grid-template-columns:160px 1fr;gap:12px;align-items:flex-start;padding:10px 8px;border-bottom:1px dashed var(--border);font-size:14px}
.row:last-child{border-bottom:none}
.badge{display:inline-flex;gap:8px;align-items:center;font-weight:600;padding:6px 10px;border-radius:10px;border:1px solid var(--border);background:#fff}
.row.ok .badge{color:#064e3b;border-color:rgba(16,185,129,.3)}
.row.warn .badge{color:#7c2d12;border-color:rgba(245,158,11,.3)}
.row.error .badge{color:#7f1d1d;border-color:rgba(239,68,68,.3)}
.row.critical .badge{color:#7f1d1d;border-color:rgba(185,28,28,.35);background:rgba(248,113,113,.07)}
.row.info .badge{color:#1e3a8a;border-color:rgba(59,130,246,.3)}
.detail{color:#1f2937;word-break:break-word}
.footer{color:var(--muted);font-size:12px;padding:8px 2px;text-align:center}
@media (max-width:700px){.row{grid-template-columns:1fr}}
'@

$script = @'
(function(){
  const input=document.getElementById("searchInput");
  if(input){
    input.addEventListener("input",function(){
      const q=this.value.toLowerCase();
      document.querySelectorAll(".row").forEach(r=>{
        r.style.display=r.textContent.toLowerCase().includes(q)?"":"none";
      });
    });
  }
  document.querySelectorAll(".section h2").forEach(h=>{
    h.addEventListener("click",()=>h.parentElement.classList.toggle("collapsed"));
  });
})();
'@

$head = @"
<!doctype html>
<html lang='fr'>
<head>
  <meta charset='utf-8'/>
  <meta name='viewport' content='width=device-width, initial-scale=1'/>
  <title>Rapport de diagnostic - $machine</title>
  <style>$styles</style>
</head>
<body>
  <div class='header'>
    <h1>Rapport de diagnostic systeme</h1>
    <div class='meta'>Genere le <strong>$(HtmlEncode($prettyTs))</strong> — Machine: <strong>$(HtmlEncode($machine))</strong> — Utilisateur: <strong>$(HtmlEncode($userId))</strong></div>
  </div>
  <div class='container'>
"@

$tools = @"
    <div class='card tools'>
      <div class='search'>
        <input id='searchInput' type='search' placeholder='Rechercher dans le rapport…'/>
      </div>
      <div class='summary'>
        $summaryHtml
      </div>
    </div>
"@

$body = @"
    $sectionsHtml
    <div class='footer'></div>
  </div>
  <script>$script</script>
</body>
</html>
"@

Set-Content -Path $htmlPath -Value ($head + $tools + $body) -Encoding UTF8

Write-Host "operation finale reussie, rapport disponible : $htmlPath"

Start-Sleep -Seconds 3
Start-Process $outputPath
