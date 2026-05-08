function Show-M365ReportData {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$InputObject,

        [Parameter()]
        [string]$Title = 'M365 Report',

        [Parameter()]
        [string]$ChartColumn,

        [Parameter()]
        [string]$ExpandColumn,

        [Parameter()]
        [switch]$ForcePopout,

        [Parameter()]
        [switch]$NoOpenBrowser,

        [Parameter()]
        [switch]$PassThru
    )

    $data = if ($null -eq $InputObject) { @() } else { @($InputObject) }
    $settings = Get-M365UiSettings

    if (($settings.BrowserPopout -eq 'None') -and (-not $ForcePopout)) {
        $data | Format-Table -AutoSize | Out-Host
        return
    }

    $jsonData = ConvertTo-Json -InputObject @($data) -Depth 8 -Compress
    if ([string]::IsNullOrWhiteSpace($jsonData)) {
        $jsonData = '[]'
    }

    $safeTitle = [System.Net.WebUtility]::HtmlEncode($Title)
    $jsSafeTitle = [string]$Title
    $jsSafeTitle = $jsSafeTitle.Replace('\\', '\\\\').Replace('"', '\\"').Replace("`r", ' ').Replace("`n", ' ')
    $safeChartColumn = if ([string]::IsNullOrWhiteSpace($ChartColumn)) { '' } else { [System.Text.RegularExpressions.Regex]::Replace($ChartColumn, '[^A-Za-z0-9_]', '') }
    $safeExpandColumn = if ([string]::IsNullOrWhiteSpace($ExpandColumn)) { '' } else { [System.Text.RegularExpressions.Regex]::Replace($ExpandColumn, '[^A-Za-z0-9_]', '') }
    $companyName = [string]$settings.CompanyName
    $logoPath = [string]$settings.LogoPath
    $htmlBrandingEnabled = [bool]$settings.HtmlBrandingEnabled
    $htmlShowCompanyName = [bool]$settings.HtmlShowCompanyName
    $htmlShowCompanyLogo = [bool]$settings.HtmlShowCompanyLogo
    $configuredSavePath = [string]$settings.ReportSavePath
    $configuredTemplate = [string]$settings.FileNameTemplate

    $safeCompanyName = if (($htmlBrandingEnabled -and $htmlShowCompanyName) -and -not [string]::IsNullOrWhiteSpace($companyName)) { [System.Net.WebUtility]::HtmlEncode($companyName) } else { '' }
    $logoUri = ''
    if (($htmlBrandingEnabled -and $htmlShowCompanyLogo) -and -not [string]::IsNullOrWhiteSpace($logoPath) -and (Test-Path -Path $logoPath)) {
      try {
        $resolvedLogoPath = Resolve-Path -Path $logoPath -ErrorAction Stop | Select-Object -ExpandProperty Path -First 1
        $logoUri = ([System.Uri]$resolvedLogoPath).AbsoluteUri
      }
      catch {
        $logoUri = ''
      }
    }

    $brandBlock = ''
    if ($htmlBrandingEnabled -and (-not [string]::IsNullOrWhiteSpace($safeCompanyName) -or -not [string]::IsNullOrWhiteSpace($logoUri))) {
      $brandLogoHtml = if ([string]::IsNullOrWhiteSpace($logoUri)) { '' } else { '<img class="brand-logo" src="' + [System.Net.WebUtility]::HtmlEncode($logoUri) + '" alt="Company logo">' }
      $brandNameHtml = if ([string]::IsNullOrWhiteSpace($safeCompanyName)) { '' } else { '<span class="brand-name">' + $safeCompanyName + '</span>' }
      $brandBlock = '<div class="brand">' + $brandLogoHtml + $brandNameHtml + '</div>'
    }

    $timeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $titleToken = ($Title -replace '[^A-Za-z0-9\-_ ]', '' -replace ' +', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($titleToken)) {
      $titleToken = 'M365-Report'
    }

    $template = if ([string]::IsNullOrWhiteSpace($configuredTemplate)) { '{Title}-{Timestamp}' } else { $configuredTemplate }
    $dateToken = Get-Date -Format 'yyyyMMdd'
    $timeToken = Get-Date -Format 'HHmmss'
    $companyToken = if ([string]::IsNullOrWhiteSpace($companyName)) { '' } else { ($companyName -replace '[^A-Za-z0-9\-_ ]', '' -replace ' +', '-').Trim('-') }

    $fileStem = $template
    $fileStem = $fileStem.Replace('{Title}', $titleToken)
    $fileStem = $fileStem.Replace('{Timestamp}', $timeStamp)
    $fileStem = $fileStem.Replace('{Date}', $dateToken)
    $fileStem = $fileStem.Replace('{Time}', $timeToken)
    $fileStem = $fileStem.Replace('{CompanyName}', $companyToken)
    $fileStem = ($fileStem -replace '[^A-Za-z0-9\-_ ]', '' -replace ' +', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($fileStem)) {
      $fileStem = "M365-Report-$timeStamp"
    }

    $reportDirectory = if (-not [string]::IsNullOrWhiteSpace($configuredSavePath)) { $configuredSavePath } else { Join-Path -Path $env:TEMP -ChildPath 'M365-Exchange-Reports' }
    if (-not (Test-Path -Path $reportDirectory)) {
        New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
    }

    $fileName = "$fileStem.html"

    $reportPath = Join-Path -Path $reportDirectory -ChildPath $fileName

    $htmlTemplate = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__HTML_TITLE__</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --primary:#0f766e;--primary-h:#0d9488;--primary-dk:#0d5c57;
  --surf:#fff;--surf2:#f8fafc;--surf3:#f1f5f9;
  --bd:#e2e8f0;--bd2:#cbd5e1;
  --t1:#0f172a;--t2:#334155;--t3:#64748b;--t4:#94a3b8;
  --green:#16a34a;--green-bg:#dcfce7;
  --red:#dc2626;--red-bg:#fee2e2;
  --amber:#b45309;--amber-bg:#fef3c7;
  --blue:#2563eb;--blue-bg:#dbeafe;
  --purple:#7c3aed;--purple-bg:#ede9fe;
  --r:8px;
  --sh:0 1px 3px rgba(0,0,0,.08),0 1px 2px rgba(0,0,0,.05);
  --sh2:0 4px 6px -1px rgba(0,0,0,.07),0 2px 4px -2px rgba(0,0,0,.05);
}
html,body{height:100%}
body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:var(--surf2);color:var(--t1);display:flex;flex-direction:column;min-height:100vh;font-size:13px;line-height:1.5}

/* topbar */
.topbar{background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);height:50px;padding:0 18px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;box-shadow:0 2px 8px rgba(0,0,0,.28)}
.tb-left{display:flex;align-items:center;gap:10px;overflow:hidden}
.tb-divider{width:1px;height:20px;background:#334155;flex-shrink:0}
.brand-logo{height:26px;width:auto;max-width:130px;object-fit:contain;border-radius:3px;flex-shrink:0}
.brand-name{color:#e2e8f0;font-size:13px;font-weight:600;white-space:nowrap;flex-shrink:0}
.tb-report-title{color:#64748b;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tb-right{display:flex;align-items:center;gap:10px;flex-shrink:0}
.tb-time{color:#475569;font-size:11px;white-space:nowrap}

/* page */
.page{flex:1;display:flex;flex-direction:column;padding:16px 18px 18px;gap:12px;min-height:0;overflow:hidden}

/* title row */
.title-row{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap;flex-shrink:0}
.title-block h1{font-size:21px;font-weight:700;color:var(--t1);letter-spacing:-.025em;line-height:1.2}
.title-meta{margin-top:3px;font-size:11px;color:var(--t3)}
.actions{display:flex;gap:7px;align-items:center;padding-top:4px}

/* buttons */
.btn{display:inline-flex;align-items:center;gap:5px;padding:7px 13px;border-radius:var(--r);font-size:12px;font-weight:500;cursor:pointer;border:1px solid transparent;transition:all .15s;white-space:nowrap;font-family:inherit}
.btn svg{width:13px;height:13px;flex-shrink:0}
.btn-primary{background:var(--primary);color:#fff;border-color:var(--primary)}
.btn-primary:hover{background:var(--primary-h);box-shadow:var(--sh2)}
.btn-ghost{background:var(--surf);color:var(--t2);border-color:var(--bd2)}
.btn-ghost:hover{background:var(--surf3)}

/* stat cards */
.stat-row{display:flex;gap:8px;flex-wrap:wrap;flex-shrink:0}
.stat-card{background:var(--surf);border:1px solid var(--bd);border-radius:var(--r);padding:9px 15px;display:flex;flex-direction:column;gap:1px;box-shadow:var(--sh);min-width:90px}
.stat-v{font-size:19px;font-weight:700;color:var(--primary);line-height:1.1;font-variant-numeric:tabular-nums}
.stat-l{font-size:10px;color:var(--t3);font-weight:600;text-transform:uppercase;letter-spacing:.05em}

/* filter bar */
.fbar{background:var(--surf);border:1px solid var(--bd);border-radius:var(--r);padding:9px 13px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;box-shadow:var(--sh);flex-shrink:0}
.flabel{font-size:10px;font-weight:700;color:var(--t4);text-transform:uppercase;letter-spacing:.06em;white-space:nowrap}
.pills{display:flex;gap:4px;flex-wrap:wrap}
.pill{padding:4px 11px;border-radius:20px;font-size:11px;font-weight:500;border:1px solid var(--bd2);background:var(--surf2);color:var(--t2);cursor:pointer;transition:all .14s;white-space:nowrap;font-family:inherit}
.pill:hover{border-color:var(--primary);color:var(--primary);background:#f0fdfa}
.pill.on{background:var(--primary);color:#fff;border-color:var(--primary)}
.fdivider{width:1px;height:22px;background:var(--bd);flex-shrink:0}
.srch{position:relative;flex:1;min-width:160px;max-width:280px}
.srch-ico{position:absolute;left:9px;top:50%;transform:translateY(-50%);width:13px;height:13px;color:var(--t4);pointer-events:none}
.srch input{width:100%;padding:6px 28px 6px 28px;border:1px solid var(--bd2);border-radius:20px;font-size:12px;background:var(--surf2);color:var(--t1);outline:none;transition:border-color .15s,box-shadow .15s;font-family:inherit}
.srch input:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(15,118,110,.12);background:#fff}
.srch-clr{position:absolute;right:8px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:var(--t4);font-size:13px;line-height:1;padding:2px;display:none}
.srch-clr:hover{color:var(--t2)}
.flt-sel{min-width:170px;padding:6px 9px;border:1px solid var(--bd2);border-radius:8px;font-size:12px;background:var(--surf2);color:var(--t1);outline:none;font-family:inherit}
.flt-sel:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(15,118,110,.12);background:#fff}
.flt-input{min-width:190px;max-width:260px;padding:6px 9px;border:1px solid var(--bd2);border-radius:8px;font-size:12px;background:var(--surf2);color:var(--t1);outline:none;font-family:inherit}
.flt-input:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(15,118,110,.12);background:#fff}

/* chart */
.chart-card{background:var(--surf);border:1px solid var(--bd);border-radius:var(--r);padding:14px 16px;box-shadow:var(--sh);display:none;flex-shrink:0}
.chart-card.on{display:block}
.chart-hdr{display:flex;align-items:baseline;gap:8px;margin-bottom:12px}
.chart-hdr h2{font-size:12px;font-weight:700;color:var(--t1);text-transform:uppercase;letter-spacing:.04em}
.chart-hdr span{font-size:11px;color:var(--t3)}
.cbar{display:flex;align-items:center;margin-bottom:4px}
.cbar-lbl{width:195px;text-align:right;padding-right:10px;font-size:11px;color:var(--t2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex-shrink:0}
.cbar-track{flex:1;background:var(--surf3);border-radius:3px;height:13px;overflow:hidden}
.cbar-fill{height:100%;border-radius:3px;background:linear-gradient(90deg,var(--primary-dk) 0%,var(--primary) 60%,#14b8a6 100%);transition:width .45s cubic-bezier(.4,0,.2,1)}
.cbar-val{width:95px;padding-left:9px;font-size:11px;color:var(--t3);font-variant-numeric:tabular-nums;flex-shrink:0}

/* table container */
.tbl-wrap{flex:1;min-height:0;background:var(--surf);border:1px solid var(--bd);border-radius:var(--r);box-shadow:var(--sh);overflow:auto;display:flex;flex-direction:column}
table{border-collapse:collapse;width:100%;min-width:600px}
thead tr{background:linear-gradient(180deg,var(--surf2) 0%,var(--surf3) 100%);border-bottom:2px solid var(--bd)}
thead th{position:sticky;top:0;z-index:2;background:linear-gradient(180deg,var(--surf2) 0%,var(--surf3) 100%);padding:8px 11px;text-align:left;font-size:10px;font-weight:700;color:var(--t3);white-space:nowrap;cursor:pointer;user-select:none;letter-spacing:.05em;text-transform:uppercase;border-right:1px solid var(--bd);transition:background .12s,color .12s}
thead th:last-child{border-right:none}
thead th:hover{background:#e8f5f4;color:var(--primary)}
thead th.sa{color:var(--primary);background:#f0fdfa}
.th-in{display:flex;align-items:center;gap:4px}
.sico{font-size:9px;color:var(--t4);flex-shrink:0}
thead th.sa .sico{color:var(--primary)}
tbody tr{border-bottom:1px solid var(--bd);transition:background .08s}
tbody tr:last-child{border-bottom:none}
tbody tr:nth-child(even){background:#fafcff}
tbody tr:hover{background:#f0fdfa}
td{padding:7px 11px;font-size:12px;color:var(--t2);white-space:normal;max-width:none;overflow:visible;text-overflow:clip;border-right:1px solid transparent;word-break:break-word}
td:last-child{border-right:none}
td.null{color:var(--t4)}

/* chips */
.chip{display:inline-flex;align-items:center;gap:3px;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:500;line-height:1.5}
.cg{background:var(--green-bg);color:var(--green)}
.cr{background:var(--red-bg);color:var(--red)}
.ca{background:var(--amber-bg);color:var(--amber)}
.cb{background:var(--blue-bg);color:var(--blue)}
.cp{background:var(--purple-bg);color:var(--purple)}
.cx{background:var(--surf3);color:var(--t3)}

/* empty state */
.empty-state{display:none;flex-direction:column;align-items:center;justify-content:center;padding:56px 20px;color:var(--t4);gap:8px;flex:1}
.empty-state svg{width:44px;height:44px;opacity:.3}
.empty-state p{font-size:14px;font-weight:600}
.empty-state small{font-size:12px}

/* footer */
.ftr{display:flex;align-items:center;justify-content:space-between;padding:5px 18px;background:var(--surf3);border-top:1px solid var(--bd);flex-shrink:0}
.ftr-l,.ftr-r{font-size:11px;color:var(--t3)}
/* hover card */
#hoverCard{position:fixed;z-index:9999;background:#fff;border:1px solid var(--bd2);border-radius:10px;box-shadow:0 8px 32px rgba(0,0,0,.18),0 2px 8px rgba(0,0,0,.1);min-width:230px;max-width:340px;pointer-events:all;opacity:0;transition:opacity .13s ease;overflow:hidden}
#hoverCard.show{opacity:1}
.hc-head{background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);padding:11px 15px 10px;color:#f1f5f9;font-size:13px;font-weight:700;word-break:break-word;line-height:1.3}
.hc-body{padding:8px 14px 10px;display:flex;flex-direction:column;gap:0;max-height:300px;overflow-y:auto}
.hc-row{display:flex;gap:8px;align-items:flex-start;padding:3px 0;border-bottom:1px solid var(--bd)}
.hc-row:last-child{border-bottom:none}
.hc-key{min-width:90px;max-width:110px;color:var(--t4);font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;flex-shrink:0;padding-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hc-val{color:var(--t2);font-size:11px;word-break:break-word;flex:1;line-height:1.4}

@media print{
  .topbar,.actions,.fbar,.chart-card,.ftr{display:none !important}
  body{background:#fff}
  .page{padding:0}
  .tbl-wrap{border:none;box-shadow:none}
  thead th{position:static}
}
</style>
</head>
<body>

<div class="topbar">
  <div class="tb-left" id="tbLeft">__BRAND_BLOCK__</div>
  <div class="tb-right"><span class="tb-time" id="tbTime"></span></div>
</div>

<div class="page">

  <div class="title-row">
    <div class="title-block">
      <h1>__HTML_TITLE__</h1>
      <div class="title-meta" id="titleMeta"></div>
    </div>
    <div class="actions">
      <button class="btn btn-primary" id="exportBtn" type="button">
        <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M8 2v8m0 0-2.5-2.5M8 10l2.5-2.5M3 13h10"/></svg>
        Export CSV
      </button>
      <button class="btn btn-ghost" id="pdfBtn" type="button">
        <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 2.5h5l3 3V13.5H4z"/><path d="M9 2.5v3h3"/><path d="M5.5 11h5"/></svg>
        Export PDF
      </button>
    </div>
  </div>

  <div class="stat-row" id="statRow"></div>

  <div class="fbar">
    <span class="flabel">View</span>
    <div class="pills" id="pills">
      <button class="pill on" data-p="all"  type="button">All</button>
      <button class="pill"    data-p="users"  type="button">Users</button>
      <button class="pill"    data-p="shared" type="button">Shared</button>
      <button class="pill"    data-p="guests" type="button">Guests</button>
      <button class="pill"    data-p="synced" type="button">On-Prem Synced</button>
    </div>
    <div class="fdivider"></div>
    <span class="flabel">Account</span>
    <div class="pills" id="acctPills">
      <button class="pill on" data-a="all"      type="button">All</button>
      <button class="pill"    data-a="enabled"  type="button">✓ Enabled</button>
      <button class="pill"    data-a="disabled" type="button">✗ Disabled</button>
    </div>
    <div class="fdivider"></div>
    <span class="flabel">Filter</span>
    <select id="fltCol" class="flt-sel"></select>
    <input id="fltVal" class="flt-input" type="text" placeholder="Column contains..." autocomplete="off" spellcheck="false">
    <div class="fdivider"></div>
    <div class="srch">
      <svg class="srch-ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="6.5" cy="6.5" r="4"/><path d="m10.5 10.5 3 3"/></svg>
      <input id="srchInput" type="text" placeholder="Search all columns…" autocomplete="off" spellcheck="false">
      <button class="srch-clr" id="srchClr" type="button" title="Clear">&#x2715;</button>
    </div>
  </div>

  <div class="chart-card" id="chartCard">
    <div class="chart-hdr"><h2 id="chartTitle"></h2><span id="chartSub"></span></div>
    <div id="chartBars"></div>
  </div>

  <div class="tbl-wrap" id="tblWrap">
    <table id="tbl"></table>
    <div class="empty-state" id="emptyState">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><rect x="3" y="3" width="18" height="18" rx="2.5"/><path d="M8 9h8M8 12h8M8 15h5"/></svg>
      <p>No matching rows</p>
      <small>Try a different profile or clear the search</small>
    </div>
  </div>

</div>

<div class="ftr">
  <span class="ftr-l" id="ftrCount"></span>
  <span class="ftr-r">M365 Reporting Tools</span>
</div>

<script>
const rowsRaw = __JSON_DATA__;
const rows = Array.isArray(rowsRaw) ? rowsRaw : (rowsRaw ? [rowsRaw] : []);
const title = "__JS_TITLE__";
const chartColumn = '__CHART_COLUMN__';
const expandColumn = '__EXPAND_COLUMN__';

const tbl    = document.getElementById('tbl');
const empty  = document.getElementById('emptyState');
const ftrCnt = document.getElementById('ftrCount');
const srch   = document.getElementById('srchInput');
const srchClr= document.getElementById('srchClr');
const fltCol = document.getElementById('fltCol');
const fltVal = document.getElementById('fltVal');

let activeRows = [...rows];
let sortState  = { col: null, dir: 'asc' };
let profile    = 'all';
let acctFilter = 'all';
let columnFilter = { col: '__all__', term: '' };

const hasProfileColumns = rows.some(r => ('UserType' in r) || ('MailboxKind' in r) || ('RecipientTypeDetails' in r) || ('OnPremisesSyncEnabled' in r));
const hasAccountColumn = rows.some(r => ('AccountEnabled' in r));

if (!hasProfileColumns) {
  const profileLabel = Array.from(document.querySelectorAll('.fbar .flabel')).find(x => (x.textContent || '').trim().toLowerCase() === 'view');
  const profilePills = document.getElementById('pills');
  if (profileLabel) { profileLabel.style.display = 'none'; }
  if (profilePills) { profilePills.style.display = 'none'; }
}

if (!hasAccountColumn) {
  const accountLabel = Array.from(document.querySelectorAll('.fbar .flabel')).find(x => (x.textContent || '').trim().toLowerCase() === 'account');
  const accountPills = document.getElementById('acctPills');
  if (accountLabel) { accountLabel.style.display = 'none'; }
  if (accountPills) { accountPills.style.display = 'none'; }
}

Array.from(document.querySelectorAll('.fbar .fdivider')).forEach(function(div){
  const prev = div.previousElementSibling;
  const next = div.nextElementSibling;
  const prevHidden = !prev || getComputedStyle(prev).display === 'none';
  const nextHidden = !next || getComputedStyle(next).display === 'none';
  if (prevHidden || nextHidden) {
    div.style.display = 'none';
  }
});

(function(){
  const colSet = new Set();
  rows.forEach(function(r){ Object.keys(r || {}).forEach(function(k){ colSet.add(k); }); });
  const cols = Array.from(colSet).sort(function(a,b){ return a.localeCompare(b, undefined, { sensitivity: 'base' }); });
  const escapedCols = cols.map(function(c){
    return c.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  });
  fltCol.innerHTML = '<option value="__all__">All columns</option>' + escapedCols.map(function(c){
    return '<option value="'+c+'">'+c+'</option>';
  }).join('');
})();

// timestamp
document.getElementById('tbTime').textContent = 'Generated ' + new Date().toLocaleString();

// stat cards
(function() {
  const row = document.getElementById('statRow');
  function card(v,l){ const d=document.createElement('div'); d.className='stat-card'; d.innerHTML='<div class="stat-v">'+v+'</div><div class="stat-l">'+l+'</div>'; return d; }
  row.appendChild(card(rows.length, 'Total Records'));
  const fcard = card(rows.length, 'Showing');
  fcard.id = 'filteredCard';
  row.appendChild(fcard);
})();

// topbar brand
(function() {
  const tbl2 = document.getElementById('tbLeft');
  const brand = tbl2.querySelector('.brand');
  if (brand) {
    const logo = brand.querySelector('.brand-logo');
    const name = brand.querySelector('.brand-name');
    brand.remove();
    if (logo) { logo.style.cssText='height:26px;width:auto;max-width:130px;object-fit:contain;border-radius:3px;flex-shrink:0'; tbl2.appendChild(logo); }
    if (name) { name.style.cssText='color:#e2e8f0;font-size:13px;font-weight:600;white-space:nowrap;flex-shrink:0'; tbl2.appendChild(name); }
    if (logo || name) { const dv=document.createElement('div'); dv.className='tb-divider'; tbl2.appendChild(dv); }
  }
  const sp = document.createElement('span'); sp.className='tb-report-title'; sp.textContent=title;
  tbl2.appendChild(sp);
})();

// helpers
function toStr(v){ if(v===null||v===undefined) return ''; if(typeof v==='object') return JSON.stringify(v); return String(v); }
function normBool(v){ if(typeof v==='boolean') return v; const t=String(v||'').toLowerCase(); return t==='true'||t==='1'||t==='yes'; }

// hover card
const _hc=(function(){
  const hc=document.createElement('div'); hc.id='hoverCard'; document.body.appendChild(hc);
  function escH(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
  function pos(x,y){
    const w=window.innerWidth,h=window.innerHeight,cw=hc.offsetWidth||270,ch=hc.offsetHeight||180;
    let l=x+16,t=y+12;
    if(l+cw>w-8){l=x-cw-12;} if(t+ch>h-8){t=h-ch-8;} if(t<8){t=8;}
    hc.style.left=l+'px'; hc.style.top=t+'px';
  }
  let _hideTimer=null;
  function cancelHide(){ if(_hideTimer){clearTimeout(_hideTimer);_hideTimer=null;} }
  function schedHide(){ cancelHide(); hc.classList.remove('show'); }
  hc.addEventListener('mouseleave',schedHide);
  document.addEventListener('keydown',function(e){ if(e.key==='Escape') schedHide(); });
  document.addEventListener('click',function(e){ if(!hc.contains(e.target)) schedHide(); });
  function show(row,x,y){
    const skip=expandColumn?[expandColumn.toLowerCase()]:[];
    const name=String(row.DisplayName||row.displayname||'');
    let body='';
    Object.keys(row).filter(function(k){
      if(k.toLowerCase()==='displayname') return false;
      if(skip.includes(k.toLowerCase())) return false;
      const v=row[k]; return v!==null&&v!==undefined&&String(v)!=='';
    }).forEach(function(k){
      const v=String(row[k]);
      const disp=v.length>70?v.substring(0,67)+'\u2026':v;
      body+='<div class="hc-row"><span class="hc-key">'+escH(k)+'</span><span class="hc-val">'+escH(disp)+'</span></div>';
    });
    hc.innerHTML='<div class="hc-head">'+escH(name||'(no name)')+'</div><div class="hc-body">'+body+'</div>';
    pos(x,y); hc.classList.add('show');
  }
  return {show:show, pos:pos};
})();

// smart cell renderer — receives full row for context-aware chips
function isSharedRow(row) {
  const ut=String(row.UserType||'').toLowerCase();
  const mk=String(row.MailboxKind||'').toLowerCase();
  const rt=String(row.RecipientTypeDetails||'').toLowerCase();
  const syn=normBool(row.OnPremisesSyncEnabled);
  const ae=normBool(row.AccountEnabled);
  const hasM=!!String(row.Mail||row.mail||row.PrimarySmtpAddress||'').trim();
  const hasSI=!!String(row.LastSuccessfulSignInDateTime||row.LastSignInDateTime||'').trim();
  return mk==='shared'||rt==='shared'||((mk==='user'||!mk)&&ut==='member'&&!ae&&hasM&&!hasSI&&!syn);
}

function renderCell(col, val, row) {
  const td = document.createElement('td');
  const raw = toStr(val);
  const cl  = col.toLowerCase();

  if (cl==='displayname') {
    td.style.cursor='context-menu';
    td.title='Right-click to view details';
    td.addEventListener('contextmenu',function(e){ e.preventDefault(); _hc.show(row,e.clientX,e.clientY); });
  }

  if (raw === '') { td.textContent = '\u2014'; td.className='null'; return td; }

  // booleans
  if (typeof val === 'boolean' || raw === 'True' || raw === 'False') {
    const bool = (val === true || raw === 'True');
    const chip = document.createElement('span');
    if (cl === 'accountenabled')        { chip.className='chip '+(bool?'cg':'cr'); chip.textContent=bool?'\u2713 Enabled':'\u2717 Disabled'; }
    else if (cl === 'passwordneverexpires') { chip.className='chip '+(bool?'ca':'cg'); chip.textContent=bool?'Never expires':'Has expiry'; }
    else if (cl === 'onpremisessyncenabled'){ chip.className='chip '+(bool?'cb':'cx'); chip.textContent=bool?'\u21c4 Synced':'Cloud only'; }
    else if (cl === 'hiddenfromaddresslistsenabled'){ chip.className='chip '+(bool?'cx':'cg'); chip.textContent=bool?'Hidden':'Visible'; }
    else { chip.className='chip '+(bool?'cg':'cx'); chip.textContent=raw; }
    td.appendChild(chip); return td;
  }

  // kind/type chips
  if (cl==='mailboxkind'||cl==='usertype'||cl==='recipienttypedetails'||cl==='grouptype'||cl==='accesstype') {
    const chip=document.createElement('span');
    // override MailboxKind 'User' to 'Shared' when row matches shared heuristic
    let display = raw;
    let rl = raw.toLowerCase();
    if (cl==='mailboxkind' && rl==='user' && row && isSharedRow(row)) { display='Shared'; rl='shared'; }
    if      (rl==='shared')                            chip.className='chip cb';
    else if (rl==='guest')                             chip.className='chip cp';
    else if (rl==='resource'||rl.includes('room')||rl.includes('equipment')) chip.className='chip ca';
    else if (rl==='user'||rl==='member'||rl==='distribution') chip.className='chip cg';
    else chip.className='chip cx';
    chip.textContent=display; td.appendChild(chip); return td;
  }

  // archive status
  if (cl==='archivestatus') {
    const chip=document.createElement('span');
    chip.className='chip '+(raw.toLowerCase()==='active'?'cg':'cx');
    chip.textContent=raw; td.appendChild(chip); return td;
  }

  // password / expiry status
  if (cl==='passwordexpirystatus') {
    const chip=document.createElement('span'); const rl=raw.toLowerCase();
    if (rl==='expired')       chip.className='chip cr';
    else if (rl==='expiring soon') chip.className='chip ca';
    else if (rl==='active')   chip.className='chip cg';
    else chip.className='chip cx';
    chip.textContent=raw; td.appendChild(chip); return td;
  }

  // dates
  if ((cl.includes('date')||cl.includes('time')||cl.includes('created')||cl.includes('logon'))&&raw.length>8) {
    try {
      const d=new Date(raw);
      if (!isNaN(d.getTime())) {
        td.textContent=d.toLocaleString(undefined,{year:'numeric',month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'});
        td.title=raw; return td;
      }
    } catch(e){}
  }

  // MB numeric
  if ((cl.includes('mb')||cl.includes('size'))&&raw!==''&&!isNaN(parseFloat(raw))) {
    td.textContent=Number(parseFloat(raw).toFixed(2)).toLocaleString()+' MB';
    td.style.fontVariantNumeric='tabular-nums'; return td;
  }

  // plain numbers
  if (!isNaN(Number(raw)) && raw.trim()!=='') {
    td.textContent=Number(raw).toLocaleString();
    td.style.fontVariantNumeric='tabular-nums'; return td;
  }

  td.textContent=raw;
  if (raw.length>70) {
    td.title=raw;
  }
  return td;
}

// sort
function sortRows(src) {
  if (!sortState.col) return src;
  const c=sortState.col, dir=sortState.dir==='asc'?1:-1;
  return src.slice().sort(function(a,b){
    const av=toStr(a[c]), bv=toStr(b[c]);
    const an=parseFloat(av), bn=parseFloat(bv);
    if (!isNaN(an)&&!isNaN(bn)) return (an-bn)*dir;
    return av.localeCompare(bv,undefined,{sensitivity:'base'})*dir;
  });
}

// expand
function expandRows(src) {
  if (!expandColumn) return src;
  const out=[];
  const groupKeys=['DisplayName','PrimarySmtpAddress','GroupType','AccessType','ManagedBy','WhenCreated','MemberCount'];
  src.forEach(function(row){
    const raw=String(row[expandColumn]||'').trim();
    if (!raw){out.push(row);return;}
    const parts=raw.split('; ').map(s=>s.trim()).filter(Boolean);
    if (parts.length<=1){out.push(row);return;}
    parts.forEach(function(p,i){
      const cp=Object.assign({},row);
      cp[expandColumn]=p;
      if (i>0) { groupKeys.forEach(k=>{ if(Object.prototype.hasOwnProperty.call(cp,k)) cp[k]=''; }); }
      out.push(cp);
    });
  });
  return out;
}

// profile
function matchProfile(row) {
  if (profile==='all') return true;
  const ut=String(row.UserType||'').toLowerCase();
  const mk=String(row.MailboxKind||'').toLowerCase();
  const rt=String(row.RecipientTypeDetails||'').toLowerCase();
  const syn=normBool(row.OnPremisesSyncEnabled);
  const ae=normBool(row.AccountEnabled);
  const hasM=!!String(row.Mail||row.mail||row.PrimarySmtpAddress||'').trim();
  const hasSI=!!String(row.LastSuccessfulSignInDateTime||row.LastSignInDateTime||'').trim();
  const byKind=mk==='shared'||rt==='shared';
  const byHeur=(mk==='user'||!mk)&&ut==='member'&&!ae&&hasM&&!hasSI&&!syn;
  const isShared=byKind||byHeur;
  if (profile==='users')  return (mk==='user'&&!isShared)||(ut==='member'&&!isShared&&mk!=='resource');
  if (profile==='shared') return isShared;
  if (profile==='guests') return ut==='guest';
  if (profile==='synced') return syn;
  return true;
}

function matchAccount(row) {
  if (acctFilter==='all') return true;
  if (!Object.prototype.hasOwnProperty.call(row,'AccountEnabled')) return true;
  const ae=normBool(row.AccountEnabled);
  if (acctFilter==='enabled')  return ae;
  if (acctFilter==='disabled') return !ae;
  return true;
}

function matchSearch(row, term) {
  if (!term) return true;
  return Object.values(row).some(v=>toStr(v).toLowerCase().includes(term));
}

function matchColumnFilter(row) {
  if (!columnFilter.term) return true;
  if (columnFilter.col === '__all__') {
    return Object.values(row).some(v=>toStr(v).toLowerCase().includes(columnFilter.term));
  }
  return toStr(row[columnFilter.col]).toLowerCase().includes(columnFilter.term);
}

// build table
function buildTable(src) {
  tbl.innerHTML='';
  empty.style.display='none';
  if (!src.length) {
    tbl.style.display='none'; empty.style.display='flex';
    ftrCnt.textContent='0 of '+rows.length+' records';
    const fc=document.getElementById('filteredCard'); if(fc) fc.querySelector('.stat-v').textContent='0';
    document.getElementById('titleMeta').textContent='No records match current filters';
    return;
  }
  tbl.style.display='table';
  const cols=Object.keys(src[0]);

  const thead=document.createElement('thead'), hr=document.createElement('tr');
  cols.forEach(col=>{
    const th=document.createElement('th');
    const inn=document.createElement('div'); inn.className='th-in';
    const lbl=document.createElement('span'); lbl.textContent=col;
    const ico=document.createElement('span'); ico.className='sico';
    if (sortState.col===col) { th.classList.add('sa'); ico.textContent=sortState.dir==='asc'?'\u25b2':'\u25bc'; }
    else ico.textContent='\u21c5';
    inn.appendChild(lbl); inn.appendChild(ico); th.appendChild(inn);
    th.addEventListener('click',function(){
      sortState.col===col?(sortState.dir=sortState.dir==='asc'?'desc':'asc'):(sortState.col=col,sortState.dir='asc');
      buildTable(sortRows(activeRows));
    });
    hr.appendChild(th);
  });
  thead.appendChild(hr);

  const tbody=document.createElement('tbody');
  src.forEach(row=>{ const tr=document.createElement('tr'); cols.forEach(c=>tr.appendChild(renderCell(c,row[c],row))); tbody.appendChild(tr); });
  tbl.appendChild(thead); tbl.appendChild(tbody);

  const cnt=src.length;
  ftrCnt.textContent=cnt+' of '+rows.length+' records';
  const fc=document.getElementById('filteredCard'); if(fc) fc.querySelector('.stat-v').textContent=cnt;
  document.getElementById('titleMeta').textContent=cnt===rows.length?rows.length+' records':cnt+' of '+rows.length+' records (filtered)';
}

// chart
function buildChart(src) {
  const card=document.getElementById('chartCard');
  if (!card||!chartColumn) return;
  const sorted=src.filter(r=>r[chartColumn]!==null&&r[chartColumn]!==undefined&&String(r[chartColumn]).trim()!=='')
    .slice().sort((a,b)=>Number(b[chartColumn])-Number(a[chartColumn])).slice(0,20);
  if (!sorted.length){ card.classList.remove('on'); return; }
  card.classList.add('on');
  const maxV=Math.max(...sorted.map(r=>Number(r[chartColumn])),0.001);
  const nk=Object.prototype.hasOwnProperty.call(sorted[0],'DisplayName')?'DisplayName':Object.keys(sorted[0])[0];
  document.getElementById('chartTitle').textContent='Top '+sorted.length+' by '+chartColumn;
  document.getElementById('chartSub').textContent='sorted largest first';
  document.getElementById('chartBars').innerHTML=sorted.map(function(row){
    const v=Number(row[chartColumn]), pct=(v/maxV*100).toFixed(1);
    const lbl=String(row[nk]||row.PrimarySmtpAddress||'').substring(0,38);
    return '<div class="cbar"><div class="cbar-lbl" title="'+lbl+'">'+lbl+'</div>'+
      '<div class="cbar-track"><div class="cbar-fill" style="width:'+pct+'%"></div></div>'+
      '<div class="cbar-val">'+v.toLocaleString(undefined,{maximumFractionDigits:2})+' MB</div></div>';
  }).join('');
}

// apply
function applyFilters() {
  const term=(srch.value||'').trim().toLowerCase();
  columnFilter.col = fltCol.value || '__all__';
  columnFilter.term = (fltVal.value || '').trim().toLowerCase();
  srchClr.style.display=term?'block':'none';
  activeRows=rows.filter(r=>matchProfile(r)&&matchAccount(r)&&matchColumnFilter(r)&&matchSearch(r,term));
  activeRows=expandRows(activeRows);
  buildTable(sortRows(activeRows));
  buildChart(activeRows);
}

// account pills
document.getElementById('acctPills').addEventListener('click',function(e){
  const btn=e.target.closest('.pill'); if(!btn) return;
  document.querySelectorAll('#acctPills .pill').forEach(p=>p.classList.remove('on'));
  btn.classList.add('on'); acctFilter=btn.dataset.a; applyFilters();
});

// profile pills
document.getElementById('pills').addEventListener('click',function(e){
  const btn=e.target.closest('.pill'); if(!btn) return;
  document.querySelectorAll('.pill').forEach(p=>p.classList.remove('on'));
  btn.classList.add('on'); profile=btn.dataset.p; applyFilters();
});

// search
srch.addEventListener('input', applyFilters);
srch.addEventListener('keydown',e=>{ if(e.key==='Escape'){srch.value='';applyFilters();} });
srchClr.addEventListener('click',()=>{ srch.value=''; srchClr.style.display='none'; applyFilters(); });
fltCol.addEventListener('change', applyFilters);
fltVal.addEventListener('input', applyFilters);
fltVal.addEventListener('keydown',e=>{ if(e.key==='Escape'){fltVal.value='';applyFilters();} });

// export
document.getElementById('exportBtn').addEventListener('click',function(){
  if (!activeRows.length) return;
  const cols=Object.keys(activeRows[0]);
  const lines=[cols.join(',')];
  activeRows.forEach(row=>lines.push(cols.map(c=>'"'+toStr(row[c]).replace(/"/g,'""')+'"').join(',')));
  const blob=new Blob([lines.join('\n')],{type:'text/csv;charset=utf-8;'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=title.replace(/\s+/g,'-')+'.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(a.href);
});

document.getElementById('pdfBtn').addEventListener('click', function(){
  window.print();
});

applyFilters();
</script>
</body>
</html>
'@

  $html = $htmlTemplate
  $html = $html.Replace('__BRAND_BLOCK__', $brandBlock)
  $html = $html.Replace('__HTML_TITLE__', $safeTitle)
  $html = $html.Replace('__JSON_DATA__', $jsonData)
  $html = $html.Replace('__JS_TITLE__', $jsSafeTitle)
  $html = $html.Replace('__CHART_COLUMN__', $safeChartColumn)
  $html = $html.Replace('__EXPAND_COLUMN__', $safeExpandColumn)

    Set-Content -Path $reportPath -Value $html -Encoding UTF8

    $fileUri = ([System.Uri]$reportPath).AbsoluteUri
    $browserLaunched = $false

    if ($NoOpenBrowser) {
      Write-Host "Generated report HTML: $reportPath" -ForegroundColor Green
      if ($PassThru) {
        return [pscustomobject]@{
          ReportPath = $reportPath
          ReportUri  = $fileUri
          Title      = $Title
          RowCount   = @($data).Count
        }
      }

      return
    }

    $selectedBrowser = if (($settings.BrowserPopout -eq 'None') -and $ForcePopout) { 'Default' } else { $settings.BrowserPopout }

    switch ($selectedBrowser) {
        'Edge' {
            $edgeCommand = Get-Command -Name msedge.exe -ErrorAction SilentlyContinue
            if ($edgeCommand) {
                Start-Process -FilePath $edgeCommand.Source -ArgumentList "--app=$fileUri"
                $browserLaunched = $true
            }
        }
        'Firefox' {
            $firefoxCommand = Get-Command -Name firefox.exe -ErrorAction SilentlyContinue
            if ($firefoxCommand) {
                Start-Process -FilePath $firefoxCommand.Source -ArgumentList "$fileUri"
                $browserLaunched = $true
            }
        }
        'Chrome' {
            $chromeCommand = Get-Command -Name chrome.exe -ErrorAction SilentlyContinue
            
            # If not in PATH, check common installation directories
            if (-not $chromeCommand) {
                $chromePaths = @(
                    'C:\Program Files\Google\Chrome\Application\chrome.exe',
                    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
                    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
                )
                foreach ($path in $chromePaths) {
                    if (Test-Path -Path $path) {
                        $chromeCommand = @{ Source = $path }
                        break
                    }
                }
            }
            
            if ($chromeCommand) {
                Start-Process -FilePath $chromeCommand.Source -ArgumentList "--app=$fileUri"
                $browserLaunched = $true
            }
        }
        'Brave' {
            $braveCommand = Get-Command -Name brave.exe -ErrorAction SilentlyContinue
            
            # If not in PATH, check common installation directories
            if (-not $braveCommand) {
                $bravePaths = @(
                    'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe',
                    'C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe',
                    "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\Application\brave.exe"
                )
                foreach ($path in $bravePaths) {
                    if (Test-Path -Path $path) {
                        $braveCommand = @{ Source = $path }
                        break
                    }
                }
            }
            
            if ($braveCommand) {
                Start-Process -FilePath $braveCommand.Source -ArgumentList "$fileUri"
                $browserLaunched = $true
            }
        }
        'Default' {
            Start-Process -FilePath $reportPath
            $browserLaunched = $true
        }
    }

    if (-not $browserLaunched -and ($selectedBrowser -ne 'None')) {
        # Fallback to default browser if selected browser not found
        Start-Process -FilePath $reportPath
    }

    Write-Host "Opened report popout: $reportPath" -ForegroundColor Green

    if ($PassThru) {
      return [pscustomobject]@{
        ReportPath = $reportPath
        ReportUri  = $fileUri
        Title      = $Title
        RowCount   = @($data).Count
      }
    }
}
