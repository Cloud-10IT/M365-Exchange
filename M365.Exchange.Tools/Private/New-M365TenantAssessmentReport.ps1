function New-M365TenantAssessmentReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Sections,

        [Parameter()]
        [string]$ReportTitle = 'Tenant Assessment Report',

        [Parameter()]
        [switch]$PassThru
    )

    $settings           = Get-M365UiSettings
    $companyName        = [string]$settings.CompanyName
    $logoPath           = [string]$settings.LogoPath
    $brandingEnabled    = [bool]$settings.HtmlBrandingEnabled
    $showName           = [bool]$settings.HtmlShowCompanyName
    $showLogo           = [bool]$settings.HtmlShowCompanyLogo
    $savePath           = if ([string]::IsNullOrWhiteSpace([string]$settings.ReportSavePath)) { $env:TEMP } else { [string]$settings.ReportSavePath }
    $fontFamily         = ([string]$settings.ReportFontFamily -replace '[^A-Za-z0-9,\- ]','').Trim()
    if ([string]::IsNullOrWhiteSpace($fontFamily)) { $fontFamily = 'Segoe UI' }

    $primary   = if (([string]$settings.ThemePrimaryColor)   -match '^#?[0-9A-Fa-f]{6}$') { if (([string]$settings.ThemePrimaryColor).StartsWith('#')) { [string]$settings.ThemePrimaryColor } else { "#$([string]$settings.ThemePrimaryColor)" } } else { '#0f766e' }
    $secondary = if (([string]$settings.ThemeSecondaryColor) -match '^#?[0-9A-Fa-f]{6}$') { if (([string]$settings.ThemeSecondaryColor).StartsWith('#')) { [string]$settings.ThemeSecondaryColor } else { "#$([string]$settings.ThemeSecondaryColor)" } } else { '#1e293b' }

    $logoUri = ''
    if ($brandingEnabled -and $showLogo -and -not [string]::IsNullOrWhiteSpace($logoPath) -and (Test-Path -Path $logoPath)) {
        try { $logoUri = ([System.Uri](Resolve-Path -Path $logoPath -ErrorAction Stop).Path).AbsoluteUri } catch {}
    }
    $safeCompany = if ($brandingEnabled -and $showName -and -not [string]::IsNullOrWhiteSpace($companyName)) { [System.Net.WebUtility]::HtmlEncode($companyName) } else { '' }

    # Build brand block
    $brandHtml = ''
    if ($logoUri)     { $brandHtml += '<img class="brand-logo" src="' + [System.Net.WebUtility]::HtmlEncode($logoUri) + '" alt="logo">' }
    if ($safeCompany) { $brandHtml += '<span class="brand-name">' + $safeCompany + '</span>' }
    if ($brandHtml)   { $brandHtml += '<div class="tb-div"></div>' }

    # Timestamp + file stem
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $titleToken = ($ReportTitle -replace '[^A-Za-z0-9\-_ ]','' -replace ' +','-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($titleToken)) { $titleToken = 'TenantAssessment' }
    $companyToken = if (-not [string]::IsNullOrWhiteSpace($companyName)) { ($companyName -replace '[^A-Za-z0-9\-_ ]','' -replace ' +','-').Trim('-') } else { '' }
    $stem = if ($companyToken) { "$companyToken-$titleToken-$timestamp" } else { "$titleToken-$timestamp" }

    if (-not (Test-Path -Path $savePath)) { New-Item -Path $savePath -ItemType Directory -Force | Out-Null }
    $reportPath = Join-Path -Path $savePath -ChildPath "$stem.html"

    # Serialize sections to JS
    $jsSections = ConvertTo-Json -InputObject @(
        $Sections | ForEach-Object {
            @{
                id          = [string]$_.Id
                title       = [string]$_.Title
                description = [string]$_.Description
                note        = if ($_.Note) { [string]$_.Note } else { '' }
                available   = [bool]$_.Available
                rows        = if ($_.Available -and $_.Rows) { @($_.Rows) } else { @() }
            }
        }
    ) -Depth 10 -Compress

    $safeTitle   = [System.Net.WebUtility]::HtmlEncode($ReportTitle)
    $jsSafeTitle = $ReportTitle.Replace('\','\\').Replace('"','\"').Replace("`r",' ').Replace("`n",' ')

    # Build tab nav HTML
    $tabNavHtml  = ($Sections | ForEach-Object { $id = [string]$_.Id; $t = [System.Net.WebUtility]::HtmlEncode([string]$_.Title); '<button class="tab" data-tab="' + $id + '" type="button">' + $t + '</button>' }) -join ''

    $html = @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>$safeTitle</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --pr:$primary;--sec:$secondary;
  --surf:#fff;--surf2:#f8fafc;--surf3:#f1f5f9;
  --bd:#e2e8f0;--bd2:#cbd5e1;
  --t1:#0f172a;--t2:#334155;--t3:#64748b;--t4:#94a3b8;
  --green:#16a34a;--gbg:#dcfce7;
  --red:#dc2626;  --rbg:#fee2e2;
  --amber:#b45309;--abg:#fef3c7;
  --blue:#2563eb; --bbg:#dbeafe;
  --purple:#7c3aed;--pbg:#ede9fe;
  --r:8px;
  --sh:0 1px 3px rgba(0,0,0,.08),0 1px 2px rgba(0,0,0,.05);
  --sh2:0 4px 6px -1px rgba(0,0,0,.07),0 2px 4px -2px rgba(0,0,0,.05);
}
html,body{height:100%;font-family:'$fontFamily',system-ui,-apple-system,sans-serif;background:var(--surf2);color:var(--t1);font-size:13px;line-height:1.5}
body{display:flex;flex-direction:column;min-height:100vh}

/* ── topbar ── */
.topbar{background:linear-gradient(135deg,var(--pr) 0%,var(--sec) 100%);padding:0 20px;height:52px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 10px rgba(0,0,0,.28);flex-shrink:0;gap:12px}
.tb-left{display:flex;align-items:center;gap:10px;overflow:hidden;flex:1;min-width:0}
.brand-logo{height:28px;width:auto;max-width:130px;object-fit:contain;border-radius:3px;flex-shrink:0}
.brand-name{color:#e2e8f0;font-size:13px;font-weight:700;white-space:nowrap;flex-shrink:0}
.tb-div{width:1px;height:20px;background:rgba(255,255,255,.2);flex-shrink:0}
.tb-title{color:rgba(255,255,255,.75);font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:500}
.tb-right{display:flex;align-items:center;gap:14px;flex-shrink:0}
.tb-time{color:rgba(255,255,255,.5);font-size:11px;white-space:nowrap}
.btn-print{display:inline-flex;align-items:center;gap:5px;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;border:1px solid rgba(255,255,255,.25);background:rgba(255,255,255,.1);color:#e2e8f0;transition:all .15s;white-space:nowrap;font-family:inherit}
.btn-print:hover{background:rgba(255,255,255,.2);border-color:rgba(255,255,255,.5)}

/* ── summary strip ── */
.summary-strip{background:var(--surf);border-bottom:1px solid var(--bd);padding:14px 20px;display:flex;gap:10px;flex-wrap:wrap;flex-shrink:0;box-shadow:var(--sh)}
.sum-card{display:flex;flex-direction:column;gap:2px;padding:8px 18px;border-radius:var(--r);border:1px solid var(--bd);background:var(--surf2);min-width:110px;flex:1;max-width:200px}
.sum-v{font-size:22px;font-weight:800;letter-spacing:-.03em;line-height:1;font-variant-numeric:tabular-nums}
.sum-l{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--t4)}
.sum-card.risk .sum-v{color:var(--red)}
.sum-card.warn .sum-v{color:var(--amber)}
.sum-card.ok   .sum-v{color:var(--green)}
.sum-card.info .sum-v{color:var(--blue)}
.sum-card.neu  .sum-v{color:var(--t2)}

/* ── tab nav ── */
.tab-nav{background:var(--surf);border-bottom:2px solid var(--bd);padding:0 20px;display:flex;gap:2px;flex-shrink:0;overflow-x:auto}
.tab{padding:12px 16px;font-size:12px;font-weight:600;border:none;background:none;cursor:pointer;color:var(--t3);border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .15s;white-space:nowrap;font-family:inherit;letter-spacing:.01em}
.tab:hover{color:var(--t1)}
.tab.active{color:var(--pr);border-bottom-color:var(--pr)}

/* ── page body ── */
.page{flex:1;display:flex;flex-direction:column;min-height:0;padding:16px 20px 20px;gap:12px;overflow:hidden}

/* ── section panel ── */
.section-panel{display:none;flex-direction:column;gap:12px;flex:1;min-height:0}
.section-panel.active{display:flex}

.sec-header{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap;flex-shrink:0}
.sec-title-block h2{font-size:18px;font-weight:700;color:var(--t1);letter-spacing:-.025em;line-height:1.2}
.sec-desc{font-size:12px;color:var(--t3);margin-top:3px}
.sec-actions{display:flex;gap:7px;align-items:center;padding-top:4px;flex-shrink:0}
.btn{display:inline-flex;align-items:center;gap:5px;padding:7px 13px;border-radius:var(--r);font-size:12px;font-weight:500;cursor:pointer;border:1px solid transparent;transition:all .15s;white-space:nowrap;font-family:inherit}
.btn-primary{background:var(--pr);color:#fff;border-color:var(--pr)}
.btn-primary:hover{opacity:.92;box-shadow:var(--sh2)}
.btn-ghost{background:var(--surf);color:var(--t2);border-color:var(--bd2)}
.btn-ghost:hover{background:var(--surf3)}

/* ── stat row ── */
.stat-row{display:flex;gap:8px;flex-wrap:wrap;flex-shrink:0}
.stat-card{background:var(--surf);border:1px solid var(--bd);border-radius:var(--r);padding:8px 14px;display:flex;flex-direction:column;gap:1px;box-shadow:var(--sh);min-width:80px}
.stat-v{font-size:18px;font-weight:700;color:var(--pr);line-height:1.1;font-variant-numeric:tabular-nums}
.stat-l{font-size:10px;color:var(--t3);font-weight:600;text-transform:uppercase;letter-spacing:.05em}

/* ── callout ── */
.callout{padding:9px 14px;border-radius:var(--r);font-size:12px;font-weight:500;border-left:3px solid;flex-shrink:0}
.callout-risk{background:var(--rbg);color:var(--red);border-color:var(--red)}
.callout-warn{background:var(--abg);color:var(--amber);border-color:var(--amber)}
.callout-ok  {background:var(--gbg);color:var(--green);border-color:var(--green)}
.callout-na  {background:var(--surf3);color:var(--t3);border-color:var(--bd2)}

/* ── filter bar ── */
.fbar{background:var(--surf);border:1px solid var(--bd);border-radius:var(--r);padding:9px 13px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;box-shadow:var(--sh);flex-shrink:0}
.flabel{font-size:10px;font-weight:700;color:var(--t4);text-transform:uppercase;letter-spacing:.06em;white-space:nowrap}
.srch{position:relative;flex:1;min-width:180px;max-width:320px}
.srch-ico{position:absolute;left:9px;top:50%;transform:translateY(-50%);width:13px;height:13px;color:var(--t4);pointer-events:none}
.srch input{width:100%;padding:6px 30px 6px 29px;border:1px solid var(--bd2);border-radius:20px;font-size:12px;background:var(--surf2);color:var(--t1);outline:none;transition:border-color .15s;font-family:inherit}
.srch input:focus{border-color:var(--pr);box-shadow:0 0 0 3px rgba(15,118,110,.12);background:#fff}
.srch-clr{position:absolute;right:8px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:var(--t4);font-size:13px;padding:2px;display:none}
.flt-sel{min-width:160px;padding:6px 8px;border:1px solid var(--bd2);border-radius:8px;font-size:12px;background:var(--surf2);color:var(--t1);outline:none;font-family:inherit}
.flt-sel:focus{border-color:var(--pr);box-shadow:0 0 0 3px rgba(15,118,110,.12)}
.flt-input{min-width:170px;max-width:240px;padding:6px 8px;border:1px solid var(--bd2);border-radius:8px;font-size:12px;background:var(--surf2);color:var(--t1);outline:none;font-family:inherit}
.flt-input:focus{border-color:var(--pr);box-shadow:0 0 0 3px rgba(15,118,110,.12)}
.fdivider{width:1px;height:22px;background:var(--bd);flex-shrink:0}
.row-count{font-size:11px;color:var(--t3);white-space:nowrap;margin-left:auto}

/* ── table ── */
.tbl-wrap{flex:1;min-height:0;background:var(--surf);border:1px solid var(--bd);border-radius:var(--r);box-shadow:var(--sh);overflow:auto;display:flex;flex-direction:column}
table{border-collapse:collapse;width:100%;min-width:600px}
thead tr{background:linear-gradient(180deg,var(--surf2) 0%,var(--surf3) 100%);border-bottom:2px solid var(--bd)}
thead th{position:sticky;top:0;z-index:2;background:linear-gradient(180deg,var(--surf2) 0%,var(--surf3) 100%);padding:8px 11px;text-align:left;font-size:10px;font-weight:700;color:var(--t3);white-space:nowrap;cursor:pointer;user-select:none;letter-spacing:.05em;text-transform:uppercase;border-right:1px solid var(--bd);transition:background .12s}
thead th:last-child{border-right:none}
thead th:hover{background:#e8f5f4;color:var(--pr)}
thead th.sa{color:var(--pr);background:#f0fdfa}
.th-in{display:flex;align-items:center;gap:4px}
.sico{font-size:9px;color:var(--t4)}
thead th.sa .sico{color:var(--pr)}
tbody tr{border-bottom:1px solid var(--bd);transition:background .08s}
tbody tr:last-child{border-bottom:none}
tbody tr:nth-child(even){background:#fafcff}
tbody tr:hover{background:#f0fdfa}
td{padding:7px 11px;font-size:12px;color:var(--t2);word-break:break-word;border-right:1px solid transparent}
td:last-child{border-right:none}
td.null{color:var(--t4)}
.chip{display:inline-flex;align-items:center;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:500;line-height:1.5}
.cg{background:var(--gbg);color:var(--green)}
.cr{background:var(--rbg);color:var(--red)}
.ca{background:var(--abg);color:var(--amber)}
.cb{background:var(--bbg);color:var(--blue)}
.cp{background:var(--pbg);color:var(--purple)}
.cx{background:var(--surf3);color:var(--t3)}

/* ── unavailable panel ── */
.unavail{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:60px 20px;gap:10px;color:var(--t4);flex:1}
.unavail svg{width:48px;height:48px;opacity:.25}
.unavail p{font-size:15px;font-weight:600}
.unavail small{font-size:12px}

/* ── empty state ── */
.empty-state{display:none;flex-direction:column;align-items:center;justify-content:center;padding:48px 20px;color:var(--t4);gap:8px;flex:1}
.empty-state svg{width:40px;height:40px;opacity:.3}
.empty-state p{font-size:14px;font-weight:600}

/* ── footer ── */
.ftr{padding:6px 20px;background:var(--surf3);border-top:1px solid var(--bd);display:flex;justify-content:space-between;flex-shrink:0}
.ftr span{font-size:11px;color:var(--t3)}

@media print{
  .topbar .btn-print,.tab-nav,.fbar,.sec-actions{display:none!important}
  .tab-nav{display:none!important}
  .section-panel{display:flex!important;page-break-inside:avoid}
  body{background:#fff}
  .page{padding:0;overflow:visible}
  .tbl-wrap{border:none;box-shadow:none;overflow:visible}
  thead th{position:static}
  .summary-strip{box-shadow:none}
}
</style>
</head>
<body>

<div class="topbar">
  <div class="tb-left" id="tbLeft">
    $brandHtml
    <span class="tb-title">$safeTitle</span>
  </div>
  <div class="tb-right">
    <span class="tb-time" id="tbTime"></span>
    <button class="btn-print" onclick="window.print()" type="button">
      <svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"><rect x="4" y="11" width="8" height="4" rx="1"/><path d="M4 11V3h8v8"/><path d="M4 4H2a1 1 0 0 0-1 1v5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V5a1 1 0 0 0-1-1h-2"/></svg>
      Print
    </button>
  </div>
</div>

<div class="summary-strip" id="summaryStrip">
</div>

<div class="tab-nav" id="tabNav">
  $tabNavHtml
</div>

<div class="page" id="mainPage">
</div>

<div class="ftr">
  <span>$safeTitle</span>
  <span id="ftrTime"></span>
</div>

<script>
"use strict";
const allSections = __SECTIONS_JSON__;
const reportTitle = "$jsSafeTitle";

// ── helpers ──────────────────────────────────────────────────────
function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function toStr(v){ if(v==null) return ''; if(typeof v==='object') return JSON.stringify(v); return String(v); }
function normBool(v){ if(typeof v==='boolean') return v; const t=String(v||'').toLowerCase(); return t==='true'||t==='1'||t==='yes'; }

// ── cell renderer ─────────────────────────────────────────────────
function renderCell(col, val) {
  const td = document.createElement('td');
  const raw = toStr(val);
  const cl  = col.toLowerCase();
  if (raw === '' || raw === 'null' || raw === 'undefined') { td.textContent = '\u2014'; td.className = 'null'; return td; }

  // booleans
  if (typeof val === 'boolean' || raw === 'True' || raw === 'False') {
    const bool = (val === true || raw === 'True');
    const chip = document.createElement('span');
    if (cl === 'accountenabled')           { chip.className = 'chip '+(bool?'cg':'cr'); chip.textContent = bool?'\u2713 Enabled':'\u2717 Disabled'; }
    else if (cl === 'passwordneverexpires'){ chip.className = 'chip '+(bool?'ca':'cg'); chip.textContent = bool?'Never expires':'Has expiry'; }
    else if (cl === 'onpremisessyncenabled'){ chip.className = 'chip '+(bool?'cb':'cx'); chip.textContent = bool?'\u21c4 Synced':'Cloud only'; }
    else if (cl === 'ismanaged')           { chip.className = 'chip '+(bool?'cg':'ca'); chip.textContent = bool?'Managed':'Unmanaged'; }
    else if (cl === 'iscompliant')         { chip.className = 'chip '+(bool?'cg':'cr'); chip.textContent = bool?'Compliant':'Not compliant'; }
    else if (cl === 'staledevice')         { chip.className = 'chip '+(bool?'cr':'cg'); chip.textContent = bool?'Stale':'Active'; }
    else if (cl === 'available')           { chip.className = 'chip '+(bool?'cg':'cx'); chip.textContent = bool?'Yes':'No'; }
    else if (cl === 'delivertomailboxandforward'){ chip.className = 'chip '+(bool?'cb':'cx'); chip.textContent = bool?'Copy + Fwd':'Fwd only'; }
    else { chip.className = 'chip '+(bool?'cg':'cx'); chip.textContent = raw; }
    td.appendChild(chip); return td;
  }

  // Status: Present / Missing / Detected / Not Detected
  if (cl==='status') {
    const chip=document.createElement('span'); const rl=raw.toLowerCase();
    if      (rl==='present') chip.className='chip cg';
    else if (rl==='missing') chip.className='chip cr';
    else if (rl==='detected') chip.className='chip ca';
    else if (rl==='not detected') chip.className='chip cg';
    else chip.className='chip cx';
    chip.textContent=raw; td.appendChild(chip); return td;
  }

  // type/kind chips
  if (cl==='usertype'||cl==='recipienttypedetails'||cl==='jointype'||cl==='state'||cl==='policystate') {
    const chip = document.createElement('span');
    const rl = raw.toLowerCase();
    if      (rl==='guest')                 chip.className='chip cp';
    else if (rl==='member'||rl==='usermailbox') chip.className='chip cg';
    else if (rl==='shared'||rl==='sharedmailbox') chip.className='chip cb';
    else if (rl.includes('room')||rl.includes('equipment')) chip.className='chip ca';
    else if (rl==='enabled')              chip.className='chip cg';
    else if (rl==='disabled')             chip.className='chip cr';
    else if (rl==='enabledformsreportingonly') chip.className='chip ca';
    else chip.className='chip cx';
    chip.textContent=raw; td.appendChild(chip); return td;
  }

  // available Yes/No
  if (cl==='available') {
    const chip=document.createElement('span'); const rl=raw.toLowerCase();
    chip.className='chip '+(rl==='yes'?'cg':'cx'); chip.textContent=raw; td.appendChild(chip); return td;
  }

  // dates
  if ((cl.includes('date')||cl.includes('time')||cl.includes('created'))&&raw.length>8) {
    try { const d=new Date(raw); if(!isNaN(d.getTime())){ td.textContent=d.toLocaleString(undefined,{year:'numeric',month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'}); td.title=raw; return td; } } catch(e){}
  }

  td.textContent=raw;
  if (raw.length>80) td.title=raw;
  return td;
}

// ── table builder ─────────────────────────────────────────────────
function SectionTable(sectionId, rows) {
  let activeRows = [...rows];
  let sortState  = { col: null, dir: 'asc' };
  let searchTerm = '';
  let fltColVal  = '__all__';
  let fltTerm    = '';

  const container = document.createElement('div');
  container.style.cssText = 'display:flex;flex-direction:column;gap:12px;flex:1;min-height:0';

  // filter bar
  const fbar = document.createElement('div'); fbar.className = 'fbar';

  const fltColSel = document.createElement('select'); fltColSel.className = 'flt-sel';
  const fltInput  = document.createElement('input');  fltInput.className  = 'flt-input'; fltInput.type='text'; fltInput.placeholder='Column contains...'; fltInput.autocomplete='off';
  const fdiv      = document.createElement('div');    fdiv.className      = 'fdivider';

  const srchWrap = document.createElement('div'); srchWrap.className = 'srch';
  srchWrap.innerHTML = '<svg class="srch-ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="6.5" cy="6.5" r="4"/><path d="m10.5 10.5 3 3"/></svg>';
  const srchInput = document.createElement('input'); srchInput.type='text'; srchInput.placeholder='Search all columns\u2026'; srchInput.autocomplete='off';
  const srchClr   = document.createElement('button'); srchClr.className='srch-clr'; srchClr.type='button'; srchClr.title='Clear'; srchClr.textContent='\u2715';
  srchWrap.appendChild(srchInput); srchWrap.appendChild(srchClr);

  const rowCountSpan = document.createElement('span'); rowCountSpan.className = 'row-count';

  // populate column filter
  const colSet = new Set();
  rows.forEach(r => Object.keys(r||{}).forEach(k => colSet.add(k)));
  const cols = Array.from(colSet);
  fltColSel.innerHTML = '<option value="__all__">All columns</option>' + cols.map(c => '<option value="'+esc(c)+'">'+esc(c)+'</option>').join('');

  fbar.appendChild(fltColSel); fbar.appendChild(fltInput); fbar.appendChild(fdiv); fbar.appendChild(srchWrap); fbar.appendChild(rowCountSpan);

  // export btn in actions (appended later outside fbar)
  const exportBtn = document.createElement('button'); exportBtn.className='btn btn-ghost'; exportBtn.type='button';
  exportBtn.innerHTML='<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M8 2v8m0 0-2.5-2.5M8 10l2.5-2.5M3 13h10"/></svg> Export CSV';
  exportBtn.addEventListener('click', function(){
    if(!activeRows.length) return;
    const ks=Object.keys(activeRows[0]);
    const lines=[ks.join(',')];
    activeRows.forEach(r=>lines.push(ks.map(k=>'"'+toStr(r[k]).replace(/"/g,'""')+'"').join(',')));
    const blob=new Blob([lines.join('\n')],{type:'text/csv;charset=utf-8;'});
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
    const sec=allSections.find(s=>s.id===sectionId);
    a.download=((sec?sec.title:'export').replace(/\s+/g,'-'))+'.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(a.href);
  });

  // tbl wrap
  const tblWrap = document.createElement('div'); tblWrap.className = 'tbl-wrap';
  const tbl     = document.createElement('table');
  const emptyEl = document.createElement('div'); emptyEl.className = 'empty-state';
  emptyEl.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><rect x="3" y="3" width="18" height="18" rx="2.5"/><path d="M8 9h8M8 12h8M8 15h5"/></svg><p>No matching rows</p>';
  tblWrap.appendChild(tbl); tblWrap.appendChild(emptyEl);

  function applyFilters() {
    const term = (srchInput.value||'').trim().toLowerCase();
    fltColVal  = fltColSel.value || '__all__';
    fltTerm    = (fltInput.value||'').trim().toLowerCase();
    srchClr.style.display = term ? 'block' : 'none';
    activeRows = rows.filter(r => {
      if (fltTerm) {
        const v = fltColVal==='__all__' ? Object.values(r).some(x=>toStr(x).toLowerCase().includes(fltTerm)) : toStr(r[fltColVal]).toLowerCase().includes(fltTerm);
        if (!v) return false;
      }
      if (term) return Object.values(r).some(x=>toStr(x).toLowerCase().includes(term));
      return true;
    });
    buildTable(sortRows(activeRows));
  }

  function sortRows(src) {
    if(!sortState.col) return src;
    const c=sortState.col, dir=sortState.dir==='asc'?1:-1;
    return src.slice().sort(function(a,b){
      const av=toStr(a[c]),bv=toStr(b[c]);
      const an=parseFloat(av),bn=parseFloat(bv);
      if(!isNaN(an)&&!isNaN(bn)) return (an-bn)*dir;
      return av.localeCompare(bv,undefined,{sensitivity:'base'})*dir;
    });
  }

  function buildTable(src) {
    tbl.innerHTML='';
    emptyEl.style.display='none';
    rowCountSpan.textContent = src.length + ' of ' + rows.length + ' rows';
    if(!src.length){ tbl.style.display='none'; emptyEl.style.display='flex'; return; }
    tbl.style.display='table';
    const ks = Object.keys(src[0]);
    const thead=document.createElement('thead'),hr=document.createElement('tr');
    ks.forEach(function(col){
      const th=document.createElement('th');
      const inn=document.createElement('div'); inn.className='th-in';
      const lbl=document.createElement('span'); lbl.textContent=col;
      const ico=document.createElement('span'); ico.className='sico';
      if(sortState.col===col){th.classList.add('sa');ico.textContent=sortState.dir==='asc'?'\u25b2':'\u25bc';}else ico.textContent='\u21c5';
      inn.appendChild(lbl); inn.appendChild(ico); th.appendChild(inn);
      th.addEventListener('click',function(){
        sortState.col===col?(sortState.dir=sortState.dir==='asc'?'desc':'asc'):(sortState.col=col,sortState.dir='asc');
        buildTable(sortRows(activeRows));
      });
      hr.appendChild(th);
    });
    thead.appendChild(hr);
    const tbody=document.createElement('tbody');
    src.forEach(row=>{ const tr=document.createElement('tr'); ks.forEach(k=>tr.appendChild(renderCell(k,row[k]))); tbody.appendChild(tr); });
    tbl.appendChild(thead); tbl.appendChild(tbody);
  }

  srchInput.addEventListener('input', applyFilters);
  srchClr.addEventListener('click',()=>{ srchInput.value=''; applyFilters(); });
  srchInput.addEventListener('keydown',e=>{ if(e.key==='Escape'){srchInput.value='';applyFilters();} });
  fltColSel.addEventListener('change', applyFilters);
  fltInput.addEventListener('input', applyFilters);
  fltInput.addEventListener('keydown',e=>{ if(e.key==='Escape'){fltInput.value='';applyFilters();} });

  container.appendChild(fbar);
  container.appendChild(tblWrap);
  applyFilters();

  return { el: container, exportBtn: exportBtn };
}

// ── section stat cards ────────────────────────────────────────────
function sectionStats(sec) {
  const rows = sec.rows || [];
  const row = document.createElement('div'); row.className = 'stat-row';
  function card(v,l){ const d=document.createElement('div'); d.className='stat-card'; d.innerHTML='<div class="stat-v">'+v+'</div><div class="stat-l">'+esc(l)+'</div>'; return d; }
  row.appendChild(card(rows.length,'Total records'));
  if (sec.id==='roles') {
    const admins = rows.filter(r=>String(r.RoleName||'').toLowerCase().includes('global admin')).length;
    const disabled = rows.filter(r=>r.AccountEnabled===false||toStr(r.AccountEnabled).toLowerCase()==='false').length;
    row.appendChild(card(admins,'Global Admin members'));
    row.appendChild(card(disabled,'Disabled accounts'));
  } else if (sec.id==='forwarding') {
    const extFwd = rows.filter(r=>String(r.ForwardingSmtpAddress||'').trim()!=='').length;
    const intFwd = rows.filter(r=>String(r.ForwardingAddress||'').trim()!==''&&String(r.ForwardingSmtpAddress||'').trim()==='').length;
    row.appendChild(card(extFwd,'External forwarding'));
    row.appendChild(card(intFwd,'Internal forwarding'));
  } else if (sec.id==='devices') {
    const stale    = rows.filter(r=>r.StaleDevice===true||toStr(r.StaleDevice).toLowerCase()==='true').length;
    const unmanaged= rows.filter(r=>r.IsManaged===false||toStr(r.IsManaged).toLowerCase()==='false').length;
    row.appendChild(card(stale,'Stale devices'));
    row.appendChild(card(unmanaged,'Unmanaged devices'));
  } else if (sec.id==='ca-analysis') {
    const missing  = rows.filter(r=>toStr(r.Status).toLowerCase()==='missing').length;
    const present  = rows.filter(r=>toStr(r.Status).toLowerCase()==='present').length;
    const disabled = rows.filter(r=>toStr(r.Status).toLowerCase()==='present'&&toStr(r.PolicyState)!=='Enabled').length;
    row.appendChild(card(present,'Policies present'));
    row.appendChild(card(missing,'Policies missing'));
    row.appendChild(card(disabled,'Present but not enabled'));
  } else if (sec.id==='ca') {
    const enabled  = rows.filter(r=>toStr(r.State).toLowerCase()==='enabled').length;
    const disabled = rows.filter(r=>toStr(r.State).toLowerCase()==='disabled').length;
    const report   = rows.filter(r=>toStr(r.State).toLowerCase()==='enabledformsreportingonly').length;
    row.appendChild(card(enabled,'Enabled'));
    row.appendChild(card(disabled,'Disabled'));
    row.appendChild(card(report,'Report-only'));
  } else if (sec.id==='features') {
    const avail = rows.filter(r=>toStr(r.Available).toLowerCase()==='yes').length;
    row.appendChild(card(avail,'Features available'));
    row.appendChild(card(rows.length-avail,'Not available'));
  } else if (sec.id==='ai-apps') {
    const detected = rows.filter(r=>toStr(r.Status).toLowerCase()==='detected').length;
    const signins  = rows.reduce((sum,r)=>sum + (parseInt(toStr(r.SignInCount),10)||0),0);
    const users    = rows.reduce((sum,r)=>sum + (parseInt(toStr(r.UniqueUserCount),10)||0),0);
    row.appendChild(card(detected,'AI apps detected'));
    row.appendChild(card(signins,'Matched sign-ins'));
    row.appendChild(card(users,'User hits (sum)'));
  }
  return row;
}

// ── build summary strip ───────────────────────────────────────────
function buildSummary() {
  const strip = document.getElementById('summaryStrip');
  function card(v,l,cls){ const d=document.createElement('div'); d.className='sum-card '+cls; d.innerHTML='<div class="sum-v">'+esc(String(v))+'</div><div class="sum-l">'+esc(l)+'</div>'; return d; }

  const roles = allSections.find(s=>s.id==='roles');
  const fwd   = allSections.find(s=>s.id==='forwarding');
  const dev   = allSections.find(s=>s.id==='devices');
  const ca    = allSections.find(s=>s.id==='ca-analysis') || allSections.find(s=>s.id==='ca');
  const ai    = allSections.find(s=>s.id==='ai-apps');

  if (roles&&roles.available) {
    const admins = (roles.rows||[]).filter(r=>String(r.RoleName||'').toLowerCase().includes('global admin')).length;
    strip.appendChild(card(admins,'Global Admins', admins>4?'risk':admins>2?'warn':'ok'));
  }
  if (fwd&&fwd.available) {
    const total = (fwd.rows||[]).length;
    strip.appendChild(card(total,'Forwarding rules',total>0?'risk':'ok'));
  }
  if (dev&&dev.available) {
    const stale = (dev.rows||[]).filter(r=>r.StaleDevice===true||toStr(r.StaleDevice).toLowerCase()==='true').length;
    strip.appendChild(card(stale,'Stale devices',stale>0?'warn':'ok'));
  }
  if (ca&&ca.available) {
    const rows = ca.rows||[];
    if (ca.id==='ca-analysis') {
      const missing = rows.filter(r=>toStr(r.Status).toLowerCase()==='missing').length;
      strip.appendChild(card(missing+'/10','CA checks missing', missing>5?'risk':missing>2?'warn':missing>0?'warn':'ok'));
    } else {
      const enabled = rows.filter(r=>toStr(r.State).toLowerCase()==='enabled').length;
      strip.appendChild(card(enabled,'CA policies enabled','info'));
    }
  }
  if (ai&&ai.available) {
    const detected = (ai.rows||[]).filter(r=>toStr(r.Status).toLowerCase()==='detected').length;
    strip.appendChild(card(detected,'AI apps detected',detected>0?'warn':'ok'));
  }

  // available sections
  const available = allSections.filter(s=>s.available).length;
  strip.appendChild(card(available + '/' + allSections.length,'Sections collected','neu'));
}

// ── build section panels ──────────────────────────────────────────
function buildPanels() {
  const page = document.getElementById('mainPage');
  allSections.forEach(function(sec) {
    const panel = document.createElement('div');
    panel.className = 'section-panel';
    panel.id = 'panel-' + sec.id;

    if (!sec.available) {
      panel.innerHTML = '<div class="sec-header"><div class="sec-title-block"><h2>'+esc(sec.title)+'</h2><div class="sec-desc">'+esc(sec.description)+'</div></div></div>'+
        '<div class="callout callout-na">'+(sec.note||'Data not collected for this section — connection not available.')+' </div>'+
        '<div class="unavail"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4m0 4h.01"/></svg><p>Section unavailable</p><small>'+(sec.note||'')+'</small></div>';
      page.appendChild(panel);
      return;
    }

    // header
    const header = document.createElement('div'); header.className = 'sec-header';
    const titleBlock = document.createElement('div'); titleBlock.className = 'sec-title-block';
    titleBlock.innerHTML = '<h2>'+esc(sec.title)+'</h2><div class="sec-desc">'+esc(sec.description)+'</div>';
    header.appendChild(titleBlock);

    if (sec.note) {
      const callout = document.createElement('div');
      callout.className = 'callout callout-' + (sec.note.toLowerCase().match(/risk|warning|critical|caution|exposed/)?'risk':sec.note.toLowerCase().match(/ok|clean|no issue|none found/)?'ok':'warn');
      callout.textContent = sec.note;
      panel.appendChild(header);
      panel.appendChild(callout);
    } else {
      panel.appendChild(header);
    }

    // stat cards
    panel.appendChild(sectionStats(sec));

    // table + actions
    if (sec.rows && sec.rows.length > 0) {
      const tblObj = SectionTable(sec.id, sec.rows);
      // attach export button to sec-header actions
      const actDiv = document.createElement('div'); actDiv.className = 'sec-actions';
      actDiv.appendChild(tblObj.exportBtn);
      titleBlock.parentElement.appendChild(actDiv);
      panel.appendChild(tblObj.el);
    } else {
      const empty = document.createElement('div'); empty.className='unavail';
      empty.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><rect x="3" y="3" width="18" height="18" rx="2.5"/><path d="M8 9h8M8 12h8M8 15h5"/></svg><p>No records found</p>';
      panel.appendChild(empty);
    }

    page.appendChild(panel);
  });
}

// ── tab switching ─────────────────────────────────────────────────
function activateTab(id) {
  document.querySelectorAll('.tab').forEach(t=>t.classList.toggle('active', t.dataset.tab===id));
  document.querySelectorAll('.section-panel').forEach(p=>p.classList.toggle('active', p.id==='panel-'+id));
}

document.querySelectorAll('.tab').forEach(function(tab){
  tab.addEventListener('click', function(){ activateTab(tab.dataset.tab); });
});

// ── init ──────────────────────────────────────────────────────────
buildSummary();
buildPanels();
const now = new Date().toLocaleString();
document.getElementById('tbTime').textContent = 'Generated ' + now;
document.getElementById('ftrTime').textContent = 'Generated ' + now;

// activate first available section
const firstSec = allSections.find(s=>s.available) || allSections[0];
if (firstSec) activateTab(firstSec.id);
</script>
</body>
</html>
"@

    $html = $html.Replace('__SECTIONS_JSON__', $jsSections)

    Set-Content -Path $reportPath -Value $html -Encoding UTF8

    if ($PassThru) {
        return [pscustomobject]@{
            ReportPath = $reportPath
            ReportUri  = ([System.Uri]$reportPath).AbsoluteUri
            Title      = $ReportTitle
        }
    }

    return $reportPath
}
