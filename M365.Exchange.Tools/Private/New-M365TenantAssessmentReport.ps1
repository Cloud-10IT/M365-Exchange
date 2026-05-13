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
    $logoDataUri        = [string]$settings.LogoDataUri
    $brandingEnabled    = [bool]$settings.HtmlBrandingEnabled
    $showName           = [bool]$settings.HtmlShowCompanyName
    $showLogo           = [bool]$settings.HtmlShowCompanyLogo
    $savePath           = if ([string]::IsNullOrWhiteSpace([string]$settings.ReportSavePath)) { $env:TEMP } else { [string]$settings.ReportSavePath }
    $fontFamily         = ([string]$settings.ReportFontFamily -replace '[^A-Za-z0-9,\- ]','').Trim()
    if ([string]::IsNullOrWhiteSpace($fontFamily)) { $fontFamily = 'Segoe UI' }

    $primary   = if (([string]$settings.ThemePrimaryColor)   -match '^#?[0-9A-Fa-f]{6}$') {
        if (([string]$settings.ThemePrimaryColor).StartsWith('#')) { [string]$settings.ThemePrimaryColor } else { "#$([string]$settings.ThemePrimaryColor)" }
    } else { '#0f766e' }

    $secondary = if (([string]$settings.ThemeSecondaryColor) -match '^#?[0-9A-Fa-f]{6}$') {
        if (([string]$settings.ThemeSecondaryColor).StartsWith('#')) { [string]$settings.ThemeSecondaryColor } else { "#$([string]$settings.ThemeSecondaryColor)" }
    } else { '#1e293b' }

    $logoUri = ''
    if ($brandingEnabled -and $showLogo) {
        if (-not [string]::IsNullOrWhiteSpace($logoDataUri) -and $logoDataUri.StartsWith('data:', [System.StringComparison]::OrdinalIgnoreCase)) {
            $logoUri = $logoDataUri
        }
        elseif (-not [string]::IsNullOrWhiteSpace($logoPath) -and (Test-Path -Path $logoPath)) {
            try {
                $logoUri = ([System.Uri](Resolve-Path -Path $logoPath -ErrorAction Stop).Path).AbsoluteUri
            }
            catch {
            }
        }
    }

    $safeCompany = if ($brandingEnabled -and $showName -and -not [string]::IsNullOrWhiteSpace($companyName)) {
        [System.Net.WebUtility]::HtmlEncode($companyName)
    } else {
        ''
    }

    # Build brand block
    $brandHtml = ''
    if ($logoUri)     { $brandHtml += '<img class="brand-logo" src="' + [System.Net.WebUtility]::HtmlEncode($logoUri) + '" alt="logo">' }
    if ($safeCompany) { $brandHtml += '<span class="brand-name">' + $safeCompany + '</span>' }
    if ($brandHtml)   { $brandHtml += '<div class="tb-div"></div>' }

    # Timestamp + file stem
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $titleToken = ($ReportTitle -replace '[^A-Za-z0-9\-_ ]','' -replace ' +','-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($titleToken)) { $titleToken = 'TenantAssessment' }

    $companyToken = if (-not [string]::IsNullOrWhiteSpace($companyName)) {
        ($companyName -replace '[^A-Za-z0-9\-_ ]','' -replace ' +','-').Trim('-')
    } else {
        ''
    }

    $stem = if ($companyToken) { "$companyToken-$titleToken-$timestamp" } else { "$titleToken-$timestamp" }

    if (-not (Test-Path -Path $savePath)) {
        New-Item -Path $savePath -ItemType Directory -Force | Out-Null
    }

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

    # Prevent HTML parsers from closing any <script> when tenant JSON/titles contain "</script>" or "<".
    # JSON.parse still decodes \u003c back to "<".
    $jsSectionsForHtml = $jsSections.Replace('<', '\u003c')
    $reportTitleJsLiteral = (ConvertTo-Json -InputObject $ReportTitle -Compress -ErrorAction Stop).Replace('<', '\u003c')

    $toolsModuleRoot   = Split-Path -Path $PSScriptRoot -Parent
    $uiSettingsPathHint = Join-Path -Path $toolsModuleRoot -ChildPath 'Config\M365.Exchange.Tools.Settings.json'
    $uiConfigBundle     = @{
        Settings         = (Get-M365UiSettings)
        SettingsFilePath = $uiSettingsPathHint
    }
    $uiConfigJsonForHtml = ($uiConfigBundle | ConvertTo-Json -Depth 8 -Compress).Replace('<', '\u003c')

    $safeTitle = [System.Net.WebUtility]::HtmlEncode($ReportTitle)

    # Build tab nav HTML
    $tabNavHtml  = (
        $Sections | ForEach-Object {
            $id = [string]$_.Id
            $t = [System.Net.WebUtility]::HtmlEncode([string]$_.Title)
            '<button class="tab" data-tab="' + $id + '" type="button">' + $t + '</button>'
        }
    ) -join ''

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

/* topbar */
.topbar{background:linear-gradient(135deg,var(--pr) 0%,var(--sec) 100%);padding:0 20px;height:52px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 10px rgba(0,0,0,.28);flex-shrink:0;gap:12px}
.tb-left{display:flex;align-items:center;gap:10px;overflow:hidden;flex:1;min-width:0}
.brand-logo{height:28px;width:auto;max-width:130px;object-fit:contain;border-radius:3px;flex-shrink:0}
.brand-name{color:#e2e8f0;font-size:13px;font-weight:700;white-space:nowrap;flex-shrink:0}
.tb-div{width:1px;height:20px;background:rgba(255,255,255,.2);flex-shrink:0}
.tb-title{color:rgba(255,255,255,.75);font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:500}
.tb-right{display:flex;align-items:center;gap:14px;flex-shrink:0}
.tb-time{color:rgba(255,255,255,.5);font-size:11px;white-space:nowrap}
.btn-print,.btn-cfg{display:inline-flex;align-items:center;gap:5px;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;border:1px solid rgba(255,255,255,.25);background:rgba(255,255,255,.1);color:#e2e8f0;transition:all .15s;white-space:nowrap;font-family:inherit}
.btn-print:hover,.btn-cfg:hover{background:rgba(255,255,255,.2);border-color:rgba(255,255,255,.5)}

/* summary strip */
.summary-strip{background:var(--surf);border-bottom:1px solid var(--bd);padding:14px 20px;display:flex;gap:10px;flex-wrap:wrap;flex-shrink:0;box-shadow:var(--sh)}
.sum-card{display:flex;flex-direction:column;gap:2px;padding:8px 18px;border-radius:var(--r);border:1px solid var(--bd);background:var(--surf2);min-width:110px;flex:1;max-width:200px}
.sum-v{display:block;font-size:22px;font-weight:800;letter-spacing:-.03em;line-height:1;font-variant-numeric:tabular-nums}
.sum-l{display:block;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--t4)}
.sum-card.risk .sum-v{color:var(--red)}
.sum-card.warn .sum-v{color:var(--amber)}
.sum-card.ok   .sum-v{color:var(--green)}
.sum-card.info .sum-v{color:var(--blue)}
.sum-card.neu  .sum-v{color:var(--t2)}
.sum-card-link{cursor:pointer;appearance:none}
.sum-card-link:hover{transform:translateY(-1px);box-shadow:var(--sh2)}
.sum-card-link:focus{outline:2px solid var(--pr);outline-offset:2px}
.sum-card:disabled{cursor:default}

/* tab nav */
.tab-nav{background:var(--surf);border-bottom:2px solid var(--bd);padding:0 20px;display:flex;gap:2px;flex-shrink:0;overflow-x:auto}
.tab{padding:12px 16px;font-size:12px;font-weight:600;border:none;background:none;cursor:pointer;color:var(--t3);border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .15s;white-space:nowrap;font-family:inherit;letter-spacing:.01em}
.tab:hover{color:var(--t1)}
.tab.active{color:var(--pr);border-bottom-color:var(--pr)}

/* page body */
.page{flex:1;display:flex;flex-direction:column;min-height:0;padding:16px 20px 20px;gap:12px;overflow:hidden}

/* section panel */
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

/* stat row */
.stat-row{display:flex;gap:8px;flex-wrap:wrap;flex-shrink:0}
.stat-card{background:var(--surf);border:1px solid var(--bd);border-radius:var(--r);padding:8px 14px;display:flex;flex-direction:column;gap:1px;box-shadow:var(--sh);min-width:80px;font:inherit;color:inherit;text-align:left}
.stat-card.stat-card-link{cursor:pointer;border-color:var(--bd2)}
.stat-card.stat-card-link:hover{border-color:var(--pr);box-shadow:var(--sh2)}
.stat-v{font-size:18px;font-weight:700;color:var(--pr);line-height:1.1;font-variant-numeric:tabular-nums;display:block}
.stat-l{font-size:10px;color:var(--t3);font-weight:600;text-transform:uppercase;letter-spacing:.05em;display:block}

/* configuration modal */
.cfg-overlay{position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:1000;display:flex;align-items:flex-start;justify-content:center;padding:24px 16px;overflow:auto}
.cfg-panel{background:var(--surf);border-radius:10px;box-shadow:var(--sh2);max-width:560px;width:100%;border:1px solid var(--bd);display:flex;flex-direction:column;max-height:calc(100vh - 48px)}
.cfg-head{padding:14px 18px;border-bottom:1px solid var(--bd);display:flex;align-items:center;justify-content:space-between;gap:10px;background:var(--surf2)}
.cfg-head h3{margin:0;font-size:15px;font-weight:700;color:var(--t1)}
.cfg-body{padding:16px 18px;overflow:auto;display:flex;flex-direction:column;gap:12px}
.cfg-row{display:flex;flex-direction:column;gap:4px}
.cfg-row label{font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.04em}
.cfg-row input,.cfg-row select{width:100%;padding:7px 10px;border:1px solid var(--bd2);border-radius:8px;font-size:13px;font-family:inherit;background:var(--surf2);color:var(--t1)}
.cfg-row .hint{font-size:11px;color:var(--t4);line-height:1.35}
.cfg-path-row{display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap}
.cfg-path-row .cfg-grow{flex:1;min-width:200px}
.cfg-path-row button{flex-shrink:0;margin-bottom:2px}
.cfg-color-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.cfg-color-row input[type=text]{flex:1;min-width:120px}
.cfg-color-row input[type=color]{width:44px;height:36px;padding:0;border:1px solid var(--bd2);border-radius:8px;cursor:pointer;background:var(--surf2)}
.cfg-foot{padding:12px 18px;border-top:1px solid var(--bd);display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;align-items:center;background:var(--surf2)}
.cfg-code{font-size:11px;word-break:break-all;background:var(--surf3);padding:8px;border-radius:6px;border:1px solid var(--bd);color:var(--t2)}

/* callout */
.callout{padding:9px 14px;border-radius:var(--r);font-size:12px;font-weight:500;border-left:3px solid;flex-shrink:0}
.callout-risk{background:var(--rbg);color:var(--red);border-color:var(--red)}
.callout-warn{background:var(--abg);color:var(--amber);border-color:var(--amber)}
.callout-ok  {background:var(--gbg);color:var(--green);border-color:var(--green)}
.callout-na  {background:var(--surf3);color:var(--t3);border-color:var(--bd2)}

/* filter bar */
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

/* table */
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

/* unavailable panel */
.unavail{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:60px 20px;gap:10px;color:var(--t4);flex:1}
.unavail svg{width:48px;height:48px;opacity:.25}
.unavail p{font-size:15px;font-weight:600}
.unavail small{font-size:12px}

/* empty state */
.empty-state{display:none;flex-direction:column;align-items:center;justify-content:center;padding:48px 20px;color:var(--t4);gap:8px;flex:1}
.empty-state svg{width:40px;height:40px;opacity:.3}
.empty-state p{font-size:14px;font-weight:600}

/* footer */
.ftr{padding:6px 20px;background:var(--surf3);border-top:1px solid var(--bd);display:flex;justify-content:space-between;flex-shrink:0}
.ftr span{font-size:11px;color:var(--t3)}

@media print{
  .topbar .btn-print,.topbar .btn-cfg,.cfg-overlay,.tab-nav,.fbar,.sec-actions{display:none!important}
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
    <button class="btn-cfg" id="btnOpenCfg" type="button" title="Edit UI settings JSON (download and replace file)">
      <svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="8" cy="8" r="2.2"/><path d="M8 1v2M8 13v2M15 8h-2M3 8H1m12.364-5.364-1.414 1.414M4.05 11.95l-1.414 1.414m0-9.192 1.414 1.414m7.9 7.9 1.414 1.414"/></svg>
      Configuration
    </button>
    <button class="btn-print" onclick="window.print()" type="button">
      <svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"><rect x="4" y="11" width="8" height="4" rx="1"/><path d="M4 11V3h8v8"/><path d="M4 4H2a1 1 0 0 0-1 1v5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V5a1 1 0 0 0-1-1h-2"/></svg>
      Print
    </button>
  </div>
</div>

<div id="cfgOverlay" class="cfg-overlay" style="display:none" aria-hidden="true">
  <div class="cfg-panel" role="dialog" aria-labelledby="cfgTitle">
    <div class="cfg-head">
      <h3 id="cfgTitle">Configuration</h3>
      <button type="button" class="btn btn-ghost" id="cfgCloseX" style="padding:4px 10px">✕</button>
    </div>
    <div class="cfg-body">
      <p class="hint" style="margin:0"><strong>Save</strong> writes <code>M365.Exchange.Tools.Settings.json</code> using the browser file picker when supported. Otherwise use <strong>Download</strong> and copy the file over the path below. <strong>Apply to this report</strong> updates colors/font on this page only (does not change the JSON file).</p>
      <div class="cfg-code" id="cfgPathHint"></div>
      <input type="file" id="cfgLogoFile" accept="image/*" style="display:none">
      <div class="cfg-row"><label for="cfgBrowserPopout">Browser popout</label><select id="cfgBrowserPopout"></select></div>
      <div class="cfg-row"><label for="cfgCompanyName">Company name</label><input id="cfgCompanyName" type="text" autocomplete="off"></div>
      <div class="cfg-path-row">
        <div class="cfg-row cfg-grow">
          <label for="cfgLogoPath">Logo path</label>
          <input id="cfgLogoPath" type="text" autocomplete="off" placeholder="File path on disk, or use Browse to embed an image">
          <span class="hint">Browse embeds the image into settings (LogoDataUri). Max ~1&nbsp;MB.</span>
        </div>
        <button type="button" class="btn btn-ghost" id="cfgBrowseLogo">Browse…</button>
      </div>
      <div class="cfg-path-row">
        <div class="cfg-row cfg-grow">
          <label for="cfgReportSavePath">Report save path</label>
          <input id="cfgReportSavePath" type="text" autocomplete="off" placeholder="Blank = user TEMP — paste full folder path from Explorer">
          <span class="hint">Browse opens a folder picker when supported; you still need to paste the full path from Explorer.</span>
        </div>
        <button type="button" class="btn btn-ghost" id="cfgBrowseReportPath">Browse…</button>
      </div>
      <div class="cfg-row">
        <label for="cfgFileNameTemplate">Title</label>
        <input id="cfgFileNameTemplate" type="text" autocomplete="off">
        <span class="hint">File name template for generated reports. Tokens: <code>{Title}</code>, <code>{Timestamp}</code></span>
      </div>
      <div class="cfg-row"><label for="cfgHtmlBrandingEnabled">HTML branding enabled</label><select id="cfgHtmlBrandingEnabled"><option value="true">true</option><option value="false">false</option></select></div>
      <div class="cfg-row"><label for="cfgHtmlShowCompanyName">Show company name in HTML</label><select id="cfgHtmlShowCompanyName"><option value="true">true</option><option value="false">false</option></select></div>
      <div class="cfg-row"><label for="cfgHtmlShowCompanyLogo">Show company logo in HTML</label><select id="cfgHtmlShowCompanyLogo"><option value="true">true</option><option value="false">false</option></select></div>
      <div class="cfg-row">
        <label for="cfgThemePrimaryColor">Theme primary (#RRGGBB)</label>
        <div class="cfg-color-row">
          <input id="cfgThemePrimaryColor" type="text" autocomplete="off" spellcheck="false">
          <input id="cfgThemePrimaryPick" type="color" value="#0f766e" title="Color picker" aria-label="Theme primary color picker">
        </div>
      </div>
      <div class="cfg-row">
        <label for="cfgThemeSecondaryColor">Theme secondary (#RRGGBB)</label>
        <div class="cfg-color-row">
          <input id="cfgThemeSecondaryColor" type="text" autocomplete="off" spellcheck="false">
          <input id="cfgThemeSecondaryPick" type="color" value="#1e293b" title="Color picker" aria-label="Theme secondary color picker">
        </div>
      </div>
      <div class="cfg-row"><label for="cfgReportFontFamily">Report font family</label><input id="cfgReportFontFamily" type="text" autocomplete="off"></div>
      <div class="cfg-row"><label for="cfgExchangeAuthMode">Exchange auth mode</label><select id="cfgExchangeAuthMode"></select></div>
    </div>
    <div class="cfg-foot">
      <button type="button" class="btn btn-ghost" id="cfgCancel">Cancel</button>
      <button type="button" class="btn btn-ghost" id="cfgApplyView">Apply to this report</button>
      <button type="button" class="btn btn-ghost" id="cfgDownload">Download M365.Exchange.Tools.Settings.json</button>
      <button type="button" class="btn btn-primary" id="cfgSave">Save</button>
    </div>
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

<script type="application/json" id="m365-assessment-sections">__SECTIONS_JSON__</script>
<script type="application/json" id="m365-ui-config">__UI_CONFIG_JSON__</script>
<script>
"use strict";
const reportTitle = $reportTitleJsLiteral;
let allSections;
try {
  allSections = JSON.parse(document.getElementById('m365-assessment-sections').textContent);
} catch (e) {
  console.error('Failed to parse assessment sections JSON', e);
  allSections = [];
}
const sectionControllers = {};

function openSectionWithPreset(sectionId, presetName) {
  const firstAvailable = allSections.find(s => s.available) || allSections[0];
  const resolvedSectionId = sectionId || (firstAvailable ? firstAvailable.id : '');
  if (!resolvedSectionId) { return; }
  activateTab(resolvedSectionId);
  const controller = sectionControllers[resolvedSectionId];
  if (controller && typeof controller.applyPreset === 'function') {
    controller.applyPreset(presetName || '');
  }
}

let uiConfigBundle = { Settings: {}, SettingsFilePath: '' };
try {
  const _cfgEl = document.getElementById('m365-ui-config');
  if (_cfgEl && _cfgEl.textContent) {
    uiConfigBundle = JSON.parse(_cfgEl.textContent);
  }
} catch (e) {
  console.error('Failed to parse embedded UI settings snapshot', e);
}

function setupSettingsModal() {
  const overlay = document.getElementById('cfgOverlay');
  const pathHint = document.getElementById('cfgPathHint');
  const bp = document.getElementById('cfgBrowserPopout');
  const ea = document.getElementById('cfgExchangeAuthMode');
  const logoFile = document.getElementById('cfgLogoFile');
  if (!overlay || !pathHint || !bp || !ea || !logoFile) return;

  let logoDataUri = '';
  const maxLogoBytes = 1048576;

  pathHint.textContent = uiConfigBundle.SettingsFilePath || '(module Config folder)';

  bp.innerHTML = ['None','Edge','Firefox','Chrome','Brave','Default'].map(function (v) {
    return '<option value="' + esc(v) + '">' + esc(v) + '</option>';
  }).join('');

  ea.innerHTML = ['Auto','Interactive','DisableWAM','Device'].map(function (v) {
    return '<option value="' + esc(v) + '">' + esc(v) + '</option>';
  }).join('');

  function readBool(id) {
    return document.getElementById(id).value === 'true';
  }

  function normalizeHex(raw) {
    var t = String(raw || '').trim();
    var m = t.match(/^#?([0-9A-Fa-f]{6})$/);
    return m ? ('#' + m[1].toUpperCase()) : '';
  }

  function syncPickersFromText() {
    var p = normalizeHex(document.getElementById('cfgThemePrimaryColor').value);
    if (p) document.getElementById('cfgThemePrimaryPick').value = p.toLowerCase();
    var s = normalizeHex(document.getElementById('cfgThemeSecondaryColor').value);
    if (s) document.getElementById('cfgThemeSecondaryPick').value = s.toLowerCase();
  }

  function wireColorPair(hexId, pickId) {
    var hex = document.getElementById(hexId);
    var pick = document.getElementById(pickId);
    pick.addEventListener('input', function () {
      hex.value = pick.value.toUpperCase();
    });
    hex.addEventListener('input', function () {
      var n = normalizeHex(hex.value);
      if (n) pick.value = n.toLowerCase();
    });
  }

  wireColorPair('cfgThemePrimaryColor', 'cfgThemePrimaryPick');
  wireColorPair('cfgThemeSecondaryColor', 'cfgThemeSecondaryPick');

  function fillForm() {
    const s = uiConfigBundle.Settings || {};
    logoDataUri = s.LogoDataUri || '';
    bp.value = s.BrowserPopout || 'Edge';
    document.getElementById('cfgCompanyName').value = s.CompanyName || '';
    document.getElementById('cfgLogoPath').value = logoDataUri ? '(Embedded image — Browse to replace, or clear this text to use Logo path only)' : (s.LogoPath || '');
    document.getElementById('cfgReportSavePath').value = s.ReportSavePath || '';
    document.getElementById('cfgFileNameTemplate').value = s.FileNameTemplate || '{Title}-{Timestamp}';
    document.getElementById('cfgHtmlBrandingEnabled').value = (s.HtmlBrandingEnabled === false) ? 'false' : 'true';
    document.getElementById('cfgHtmlShowCompanyName').value = (s.HtmlShowCompanyName === false) ? 'false' : 'true';
    document.getElementById('cfgHtmlShowCompanyLogo').value = (s.HtmlShowCompanyLogo === false) ? 'false' : 'true';
    document.getElementById('cfgThemePrimaryColor').value = s.ThemePrimaryColor || '#0f766e';
    document.getElementById('cfgThemeSecondaryColor').value = s.ThemeSecondaryColor || '#1e293b';
    document.getElementById('cfgReportFontFamily').value = s.ReportFontFamily || 'Segoe UI';
    ea.value = s.ExchangeAuthMode || 'Auto';
    syncPickersFromText();
  }

  function gatherOut() {
    var lp = document.getElementById('cfgLogoPath').value.trim();
    if (lp.indexOf('(Embedded') === 0) lp = '';
    return {
      BrowserPopout: bp.value,
      CompanyName: document.getElementById('cfgCompanyName').value.trim(),
      LogoPath: lp,
      LogoDataUri: logoDataUri || '',
      ReportSavePath: document.getElementById('cfgReportSavePath').value.trim(),
      FileNameTemplate: document.getElementById('cfgFileNameTemplate').value.trim(),
      HtmlBrandingEnabled: readBool('cfgHtmlBrandingEnabled'),
      HtmlShowCompanyName: readBool('cfgHtmlShowCompanyName'),
      HtmlShowCompanyLogo: readBool('cfgHtmlShowCompanyLogo'),
      ThemePrimaryColor: document.getElementById('cfgThemePrimaryColor').value.trim(),
      ThemeSecondaryColor: document.getElementById('cfgThemeSecondaryColor').value.trim(),
      ReportFontFamily: document.getElementById('cfgReportFontFamily').value.trim(),
      ExchangeAuthMode: ea.value
    };
  }

  function applyThemeToOpenReport() {
    var pr = normalizeHex(document.getElementById('cfgThemePrimaryColor').value);
    var sc = normalizeHex(document.getElementById('cfgThemeSecondaryColor').value);
    if (pr) document.documentElement.style.setProperty('--pr', pr);
    if (sc) document.documentElement.style.setProperty('--sec', sc);
    var ff = document.getElementById('cfgReportFontFamily').value.trim();
    if (ff) {
      var safe = ff.replace(/[^\w\s,\-]/g, '').trim() || 'Segoe UI';
      document.body.style.fontFamily = "'" + safe + "',system-ui,sans-serif";
    }
  }

  function downloadBlob(blob) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'M365.Exchange.Tools.Settings.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
  }

  function closeCfg() {
    overlay.style.display = 'none';
    overlay.setAttribute('aria-hidden', 'true');
  }

  document.getElementById('btnOpenCfg').addEventListener('click', function () {
    fillForm();
    overlay.style.display = 'flex';
    overlay.setAttribute('aria-hidden', 'false');
  });

  document.getElementById('cfgCloseX').addEventListener('click', closeCfg);
  document.getElementById('cfgCancel').addEventListener('click', closeCfg);
  overlay.addEventListener('click', function (e) {
    if (e.target === overlay) closeCfg();
  });

  document.getElementById('cfgBrowseLogo').addEventListener('click', function () {
    logoFile.value = '';
    logoFile.click();
  });

  logoFile.addEventListener('change', function (ev) {
    var f = ev.target.files && ev.target.files[0];
    if (!f) return;
    if (f.size > maxLogoBytes) {
      window.alert('Logo file must be ' + (maxLogoBytes / 1024) + ' KB or smaller.');
      return;
    }
    var reader = new FileReader();
    reader.onload = function () {
      logoDataUri = reader.result || '';
      document.getElementById('cfgLogoPath').value = '(Embedded image — Browse to replace, or clear this text to use Logo path only)';
    };
    reader.readAsDataURL(f);
  });

  document.getElementById('cfgBrowseReportPath').addEventListener('click', function () {
    if (window.showDirectoryPicker) {
      window.showDirectoryPicker().then(function () {
        window.alert('Folder selected. Browsers do not expose the full drive path. Copy the path from Windows Explorer (address bar) and paste it into Report save path.');
        document.getElementById('cfgReportSavePath').focus();
      }).catch(function (err) {
        if (err && err.name === 'AbortError') return;
        var p = window.prompt('Paste the full folder path for HTML reports (e.g. C:\\\\Reports):', document.getElementById('cfgReportSavePath').value || '');
        if (p !== null) document.getElementById('cfgReportSavePath').value = p;
      });
    } else {
      var p2 = window.prompt('Paste the full folder path for HTML reports (e.g. C:\\\\Reports):', document.getElementById('cfgReportSavePath').value || '');
      if (p2 !== null) document.getElementById('cfgReportSavePath').value = p2;
    }
  });

  document.getElementById('cfgApplyView').addEventListener('click', function () {
    applyThemeToOpenReport();
  });

  document.getElementById('cfgDownload').addEventListener('click', function () {
    var blob = new Blob([JSON.stringify(gatherOut(), null, 2)], { type: 'application/json;charset=utf-8' });
    downloadBlob(blob);
  });

  document.getElementById('cfgLogoPath').addEventListener('input', function () {
    var v = this.value.trim();
    if (v.indexOf('(Embedded') !== 0 && logoDataUri) {
      logoDataUri = '';
    }
  });

  document.getElementById('cfgSave').addEventListener('click', function () {
    var blob = new Blob([JSON.stringify(gatherOut(), null, 2)], { type: 'application/json;charset=utf-8' });
    if (window.showSaveFilePicker) {
      window.showSaveFilePicker({
        suggestedName: 'M365.Exchange.Tools.Settings.json',
        types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }]
      }).then(function (handle) {
        return handle.createWritable();
      }).then(function (writable) {
        return writable.write(blob).then(function () { return writable.close(); });
      }).then(function () {
        window.alert('Settings file saved.');
        closeCfg();
      }).catch(function (e) {
        if (e && e.name === 'AbortError') return;
        downloadBlob(blob);
        window.alert('Save picker was not available or failed in this context. A download was started instead — copy the file to the path shown above.');
      });
    } else {
      downloadBlob(blob);
      window.alert('This browser does not support Save to file here. A download was started instead.');
    }
  });
}

function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function toStr(v){ if(v==null) return ''; if(typeof v==='object') return JSON.stringify(v); return String(v); }
function normBool(v){ if(typeof v==='boolean') return v; const t=String(v||'').toLowerCase(); return t==='true'||t==='1'||t==='yes'; }

function renderCell(col, val) {
  const td = document.createElement('td');
  const raw = toStr(val);
  const cl  = col.toLowerCase();

  if (raw === '' || raw === 'null' || raw === 'undefined') {
    td.textContent = '—';
    td.className = 'null';
    return td;
  }

  // booleans
  if (typeof val === 'boolean' || raw === 'True' || raw === 'False') {
    const bool = (val === true || raw === 'True');
    const chip = document.createElement('span');

    if (cl === 'accountenabled') {
      chip.className = 'chip ' + (bool ? 'cg' : 'cr');
      chip.textContent = bool ? '✓ Enabled' : '✗ Disabled';
    }
    else if (cl === 'passwordneverexpires') {
      chip.className = 'chip ' + (bool ? 'ca' : 'cg');
      chip.textContent = bool ? 'Never expires' : 'Has expiry';
    }
    else if (cl === 'onpremisessyncenabled') {
      chip.className = 'chip ' + (bool ? 'cb' : 'cx');
      chip.textContent = bool ? '⇄ Synced' : 'Cloud only';
    }
    else if (cl === 'ismanaged') {
      chip.className = 'chip ' + (bool ? 'cg' : 'ca');
      chip.textContent = bool ? 'Managed' : 'Unmanaged';
    }
    else if (cl === 'iscompliant') {
      chip.className = 'chip ' + (bool ? 'cg' : 'cr');
      chip.textContent = bool ? 'Compliant' : 'Not compliant';
    }
    else if (cl === 'staledevice') {
      chip.className = 'chip ' + (bool ? 'cr' : 'cg');
      chip.textContent = bool ? 'Stale' : 'Active';
    }
    else if (cl === 'available') {
      chip.className = 'chip ' + (bool ? 'cg' : 'cx');
      chip.textContent = bool ? 'Yes' : 'No';
    }
    else if (cl === 'delivertomailboxandforward') {
      chip.className = 'chip ' + (bool ? 'cb' : 'cx');
      chip.textContent = bool ? 'Copy + Fwd' : 'Fwd only';
    }
    else {
      chip.className = 'chip ' + (bool ? 'cg' : 'cx');
      chip.textContent = raw;
    }

    td.appendChild(chip);
    return td;
  }

  // Status: Present / Missing / Detected / Not Detected
  if (cl === 'status') {
    const chip = document.createElement('span');
    const rl = raw.toLowerCase();

    if      (rl === 'present')      chip.className = 'chip cg';
    else if (rl === 'missing')      chip.className = 'chip cr';
    else if (rl === 'detected')     chip.className = 'chip ca';
    else if (rl === 'not detected') chip.className = 'chip cg';
    else                            chip.className = 'chip cx';

    chip.textContent = raw;
    td.appendChild(chip);
    return td;
  }

  // type/kind chips
  if (cl === 'usertype' || cl === 'recipienttypedetails' || cl === 'jointype' || cl === 'state' || cl === 'policystate') {
    const chip = document.createElement('span');
    const rl = raw.toLowerCase();

    if      (rl === 'guest')                     chip.className = 'chip cp';
    else if (rl === 'member' || rl === 'usermailbox') chip.className = 'chip cg';
    else if (rl === 'shared' || rl === 'sharedmailbox') chip.className = 'chip cb';
    else if (rl.includes('room') || rl.includes('equipment')) chip.className = 'chip ca';
    else if (rl === 'enabled')                  chip.className = 'chip cg';
    else if (rl === 'disabled')                 chip.className = 'chip cr';
    else if (rl === 'enabledformsreportingonly') chip.className = 'chip ca';
    else                                        chip.className = 'chip cx';

    chip.textContent = raw;
    td.appendChild(chip);
    return td;
  }

  // available Yes/No
  if (cl === 'available') {
    const chip = document.createElement('span');
    const rl = raw.toLowerCase();
    chip.className = 'chip ' + (rl === 'yes' ? 'cg' : 'cx');
    chip.textContent = raw;
    td.appendChild(chip);
    return td;
  }

  // dates
  if ((cl.includes('date') || cl.includes('time') || cl.includes('created')) && raw.length > 8) {
    try {
      const d = new Date(raw);
      if (!isNaN(d.getTime())) {
        td.textContent = d.toLocaleString(undefined, {
          year:'numeric',
          month:'short',
          day:'numeric',
          hour:'2-digit',
          minute:'2-digit'
        });
        td.title = raw;
        return td;
      }
    } catch(e) {}
  }

  td.textContent = raw;
  if (raw.length > 80) td.title = raw;
  return td;
}

function SectionTable(sectionId, rows) {
  let activeRows = [...rows];
  let sortState = { col: null, dir: 'asc' };
  let fltColVal = '__all__';
  let fltTerm = '';
  let statRowPredicate = null;

  const container = document.createElement('div');
  container.style.cssText = 'display:flex;flex-direction:column;gap:12px;flex:1;min-height:0';

  const fbar = document.createElement('div');
  fbar.className = 'fbar';

  const fltColSel = document.createElement('select');
  fltColSel.className = 'flt-sel';

  const fltInput = document.createElement('input');
  fltInput.className = 'flt-input';
  fltInput.type = 'text';
  fltInput.placeholder = 'Column contains...';
  fltInput.autocomplete = 'off';

  const fdiv = document.createElement('div');
  fdiv.className = 'fdivider';

  const srchWrap = document.createElement('div');
  srchWrap.className = 'srch';
  srchWrap.innerHTML = '<svg class="srch-ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="6.5" cy="6.5" r="4"/><path d="m10.5 10.5 3 3"/></svg>';

  const srchInput = document.createElement('input');
  srchInput.type = 'text';
  srchInput.placeholder = 'Search all columns…';
  srchInput.autocomplete = 'off';

  const srchClr = document.createElement('button');
  srchClr.className = 'srch-clr';
  srchClr.type = 'button';
  srchClr.title = 'Clear';
  srchClr.textContent = '✕';

  srchWrap.appendChild(srchInput);
  srchWrap.appendChild(srchClr);

  const rowCountSpan = document.createElement('span');
  rowCountSpan.className = 'row-count';

  const colSet = new Set();
  rows.forEach(r => Object.keys(r || {}).forEach(k => colSet.add(k)));
  const cols = Array.from(colSet);

  fltColSel.innerHTML =
    '<option value="__all__">All columns</option>' +
    cols.map(c => '<option value="' + esc(c) + '">' + esc(c) + '</option>').join('');

  fbar.appendChild(fltColSel);
  fbar.appendChild(fltInput);
  fbar.appendChild(fdiv);
  fbar.appendChild(srchWrap);
  fbar.appendChild(rowCountSpan);

  const exportBtn = document.createElement('button');
  exportBtn.className = 'btn btn-ghost';
  exportBtn.type = 'button';
  exportBtn.innerHTML = '<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M8 2v8m0 0-2.5-2.5M8 10l2.5-2.5M3 13h10"/></svg> Export CSV';
  exportBtn.addEventListener('click', function () {
    if (!activeRows.length) return;
    const ks = Object.keys(activeRows[0]);
    const lines = [ks.join(',')];
    activeRows.forEach(r => lines.push(ks.map(k => '"' + toStr(r[k]).replace(/"/g, '""') + '"').join(',')));
    const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    const sec = allSections.find(s => s.id === sectionId);
    a.download = ((sec ? sec.title : 'export').replace(/\s+/g, '-')) + '.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
  });

  const tblWrap = document.createElement('div');
  tblWrap.className = 'tbl-wrap';

  const tbl = document.createElement('table');

  const emptyEl = document.createElement('div');
  emptyEl.className = 'empty-state';
  emptyEl.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><rect x="3" y="3" width="18" height="18" rx="2.5"/><path d="M8 9h8M8 12h8M8 15h5"/></svg><p>No matching rows</p>';

  tblWrap.appendChild(tbl);
  tblWrap.appendChild(emptyEl);

  function applyFilters() {
    const term = (srchInput.value || '').trim().toLowerCase();
    fltColVal = fltColSel.value || '__all__';
    fltTerm = (fltInput.value || '').trim().toLowerCase();
    srchClr.style.display = term ? 'block' : 'none';

    activeRows = rows.filter(r => {
      const rv = r && typeof r === 'object' ? r : {};
      if (statRowPredicate && !statRowPredicate(rv)) return false;

      if (fltTerm) {
        let matchesColumn;
        if (fltColVal === '__all__') {
          matchesColumn = Object.keys(rv).some(k => toStr(rv[k]).toLowerCase().includes(fltTerm));
        } else if (fltTerm === '__nonempty__') {
          matchesColumn = toStr(rv[fltColVal]).trim() !== '';
        } else {
          matchesColumn = toStr(rv[fltColVal]).toLowerCase().includes(fltTerm);
        }
        if (!matchesColumn) return false;
      }

      if (term) {
        return Object.keys(rv).some(k => toStr(rv[k]).toLowerCase().includes(term));
      }

      return true;
    });

    buildTable(sortRows(activeRows));
  }

  function sortRows(src) {
    if (!sortState.col) return src;

    const c = sortState.col;
    const dir = sortState.dir === 'asc' ? 1 : -1;

    return src.slice().sort(function (a, b) {
      const av = toStr(a[c]);
      const bv = toStr(b[c]);
      const an = parseFloat(av);
      const bn = parseFloat(bv);

      if (!isNaN(an) && !isNaN(bn)) {
        return (an - bn) * dir;
      }

      return av.localeCompare(bv, undefined, { sensitivity: 'base' }) * dir;
    });
  }

  function buildTable(src) {
    tbl.innerHTML = '';
    emptyEl.style.display = 'none';
    rowCountSpan.textContent = src.length + ' of ' + rows.length + ' rows';

    if (!src.length) {
      tbl.style.display = 'none';
      emptyEl.style.display = 'flex';
      return;
    }

    tbl.style.display = 'table';

    const ks = Object.keys(src[0]);
    const thead = document.createElement('thead');
    const hr = document.createElement('tr');

    ks.forEach(function (col) {
      const th = document.createElement('th');
      const inn = document.createElement('div');
      inn.className = 'th-in';

      const lbl = document.createElement('span');
      lbl.textContent = col;

      const ico = document.createElement('span');
      ico.className = 'sico';

      if (sortState.col === col) {
        th.classList.add('sa');
        ico.textContent = sortState.dir === 'asc' ? '▲' : '▼';
      } else {
        ico.textContent = '⇅';
      }

      inn.appendChild(lbl);
      inn.appendChild(ico);
      th.appendChild(inn);

      th.addEventListener('click', function () {
        if (sortState.col === col) {
          sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
        } else {
          sortState.col = col;
          sortState.dir = 'asc';
        }

        buildTable(sortRows(activeRows));
      });

      hr.appendChild(th);
    });

    thead.appendChild(hr);

    const tbody = document.createElement('tbody');
    src.forEach(row => {
      const tr = document.createElement('tr');
      ks.forEach(k => tr.appendChild(renderCell(k, row[k])));
      tbody.appendChild(tr);
    });

    tbl.appendChild(thead);
    tbl.appendChild(tbody);
  }

  function resetFilters() {
    statRowPredicate = null;
    fltColSel.value = '__all__';
    fltInput.value = '';
    srchInput.value = '';
    srchClr.style.display = 'none';
  }

  function applyStatKey(idx) {
    resetFilters();
    if (idx === 0) {
      applyFilters();
      return;
    }

    switch (sectionId) {
      case 'roles':
        if (idx === 1) { fltColSel.value = 'RoleName'; fltInput.value = 'Global Administrator'; }
        else if (idx === 2) { fltColSel.value = 'AccountEnabled'; fltInput.value = 'false'; }
        break;
      case 'forwarding':
        if (idx === 1) {
          statRowPredicate = function (r) { return toStr(r.ForwardingSmtpAddress).indexOf('@') >= 0; };
        } else if (idx === 2) {
          statRowPredicate = function (r) {
            return toStr(r.ForwardingSmtpAddress).indexOf('@') < 0 && toStr(r.ForwardingAddress).trim() !== '';
          };
        }
        break;
      case 'devices':
        if (idx === 1) { fltColSel.value = 'StaleDevice'; fltInput.value = 'True'; }
        else if (idx === 2) { fltColSel.value = 'IsManaged'; fltInput.value = 'False'; }
        break;
      case 'ca-analysis':
        if (idx === 1) { fltColSel.value = 'Status'; fltInput.value = 'Present'; }
        else if (idx === 2) { fltColSel.value = 'Status'; fltInput.value = 'Missing'; }
        else if (idx === 3) {
          statRowPredicate = function (r) {
            return toStr(r.Status).toLowerCase() === 'present' && toStr(r.PolicyState).toLowerCase() !== 'enabled';
          };
        }
        break;
      case 'ca':
        if (idx === 1) { fltColSel.value = 'State'; fltInput.value = 'Enabled'; }
        else if (idx === 2) { fltColSel.value = 'State'; fltInput.value = 'Disabled'; }
        else if (idx === 3) { fltColSel.value = 'State'; fltInput.value = 'Reporting'; }
        break;
      case 'features':
        if (idx === 1) { fltColSel.value = 'Available'; fltInput.value = 'Yes'; }
        else if (idx === 2) { fltColSel.value = 'Available'; fltInput.value = 'No'; }
        break;
      case 'ai-apps':
        if (idx === 1) { fltColSel.value = 'Status'; fltInput.value = 'Detected'; }
        else if (idx === 2) { fltColSel.value = 'SignInCount'; fltInput.value = '__nonempty__'; }
        else if (idx === 3) { fltColSel.value = 'UniqueUserCount'; fltInput.value = '__nonempty__'; }
        break;
      default:
        break;
    }

    applyFilters();
  }

  function applyPreset(presetName) {
    resetFilters();

    switch (sectionId) {
      case 'roles':
        if (presetName === 'global-admins') {
          fltColSel.value = 'RoleName';
          fltInput.value = 'Global Administrator';
        }
        break;

      case 'forwarding':
        if (presetName === 'has-forwarding') {
          statRowPredicate = function (r) {
            return toStr(r.ForwardingSmtpAddress).indexOf('@') >= 0 ||
              (toStr(r.ForwardingSmtpAddress).indexOf('@') < 0 && toStr(r.ForwardingAddress).trim() !== '');
          };
        }
        break;

      case 'devices':
        if (presetName === 'stale-devices') {
          fltColSel.value = 'StaleDevice';
          fltInput.value = 'True';
        }
        break;

      case 'ca-analysis':
        if (presetName === 'ca-missing') {
          fltColSel.value = 'Status';
          fltInput.value = 'Missing';
        }
        break;

      case 'ai-apps':
        if (presetName === 'ai-detected') {
          fltColSel.value = 'Status';
          fltInput.value = 'Detected';
        }
        break;

      case 'features':
        if (presetName === 'features-available') {
          fltColSel.value = 'Available';
          fltInput.value = 'Yes';
        }
        else if (presetName === 'features-unavailable') {
          fltColSel.value = 'Available';
          fltInput.value = 'No';
        }
        break;
    }

    applyFilters();
  }

  srchInput.addEventListener('input', applyFilters);
  srchClr.addEventListener('click', () => {
    srchInput.value = '';
    applyFilters();
  });
  srchInput.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      srchInput.value = '';
      applyFilters();
    }
  });

  fltColSel.addEventListener('change', applyFilters);
  fltInput.addEventListener('input', applyFilters);
  fltInput.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      fltInput.value = '';
      applyFilters();
    }
  });

  container.appendChild(fbar);
  container.appendChild(tblWrap);

  applyFilters();

  return {
    el: container,
    exportBtn: exportBtn,
    applyPreset: applyPreset,
    applyStatKey: applyStatKey
  };
}

function sectionStats(sec, tableCtl) {
  const rows = sec.rows || [];
  const row = document.createElement('div');
  row.className = 'stat-row';

  function makeStatCard(v, l, idx) {
    const clickable = !!(tableCtl && typeof tableCtl.applyStatKey === 'function');
    const el = document.createElement(clickable ? 'button' : 'div');
    if (clickable) el.type = 'button';
    el.className = 'stat-card' + (clickable ? ' stat-card-link' : '');
    el.title = clickable ? ('Filter: ' + l) : '';
    el.innerHTML = '<span class="stat-v">' + esc(String(v)) + '</span><span class="stat-l">' + esc(l) + '</span>';
    if (clickable) {
      el.addEventListener('click', function () { tableCtl.applyStatKey(idx); });
    }
    return el;
  }

  row.appendChild(makeStatCard(rows.length, 'Total records', 0));

  if (sec.id === 'roles') {
    const admins = rows.filter(r => String(r.RoleName || '').toLowerCase().includes('global admin')).length;
    const disabled = rows.filter(r => r.AccountEnabled === false || toStr(r.AccountEnabled).toLowerCase() === 'false').length;
    row.appendChild(makeStatCard(admins, 'Global Admin members', 1));
    row.appendChild(makeStatCard(disabled, 'Disabled accounts', 2));
  }
  else if (sec.id === 'forwarding') {
    const extFwd = rows.filter(r => String(r.ForwardingSmtpAddress || '').trim() !== '').length;
    const intFwd = rows.filter(r => String(r.ForwardingAddress || '').trim() !== '' && String(r.ForwardingSmtpAddress || '').trim() === '').length;
    row.appendChild(makeStatCard(extFwd, 'External forwarding', 1));
    row.appendChild(makeStatCard(intFwd, 'Internal forwarding', 2));
  }
  else if (sec.id === 'devices') {
    const stale = rows.filter(r => r.StaleDevice === true || toStr(r.StaleDevice).toLowerCase() === 'true').length;
    const unmanaged = rows.filter(r => r.IsManaged === false || toStr(r.IsManaged).toLowerCase() === 'false').length;
    row.appendChild(makeStatCard(stale, 'Stale devices', 1));
    row.appendChild(makeStatCard(unmanaged, 'Unmanaged devices', 2));
  }
  else if (sec.id === 'ca-analysis') {
    const missing = rows.filter(r => toStr(r.Status).toLowerCase() === 'missing').length;
    const present = rows.filter(r => toStr(r.Status).toLowerCase() === 'present').length;
    const disabled = rows.filter(r => toStr(r.Status).toLowerCase() === 'present' && toStr(r.PolicyState) !== 'Enabled').length;
    row.appendChild(makeStatCard(present, 'Policies present', 1));
    row.appendChild(makeStatCard(missing, 'Policies missing', 2));
    row.appendChild(makeStatCard(disabled, 'Present but not enabled', 3));
  }
  else if (sec.id === 'ca') {
    const enabled = rows.filter(r => toStr(r.State).toLowerCase() === 'enabled').length;
    const disabled = rows.filter(r => toStr(r.State).toLowerCase() === 'disabled').length;
    const report = rows.filter(r => toStr(r.State).toLowerCase() === 'enabledformsreportingonly').length;
    row.appendChild(makeStatCard(enabled, 'Enabled', 1));
    row.appendChild(makeStatCard(disabled, 'Disabled', 2));
    row.appendChild(makeStatCard(report, 'Report-only', 3));
  }
  else if (sec.id === 'features') {
    const avail = rows.filter(r => toStr(r.Available).toLowerCase() === 'yes').length;
    row.appendChild(makeStatCard(avail, 'Features available', 1));
    row.appendChild(makeStatCard(rows.length - avail, 'Not available', 2));
  }
  else if (sec.id === 'ai-apps') {
    const detected = rows.filter(r => toStr(r.Status).toLowerCase() === 'detected').length;
    const signins = rows.reduce((sum,r) => sum + (parseInt(toStr(r.SignInCount),10) || 0), 0);
    const users = rows.reduce((sum,r) => sum + (parseInt(toStr(r.UniqueUserCount),10) || 0), 0);
    row.appendChild(makeStatCard(detected, 'AI apps detected', 1));
    row.appendChild(makeStatCard(signins, 'Matched sign-ins', 2));
    row.appendChild(makeStatCard(users, 'User hits (sum)', 3));
  }

  return row;
}

function buildSummary() {
  const strip = document.getElementById('summaryStrip');

  function card(v, l, cls, onClick) {
    const d = document.createElement(typeof onClick === 'function' ? 'button' : 'div');

    if (d.tagName === 'BUTTON') {
      d.type = 'button';
    }

    d.className = 'sum-card ' + cls;
    d.innerHTML = '<span class="sum-v">' + esc(String(v)) + '</span><span class="sum-l">' + esc(String(l)) + '</span>';

    if (typeof onClick === 'function') {
      d.classList.add('sum-card-link');
      d.addEventListener('click', onClick);
    }

    return d;
  }

  const roles = allSections.find(s => s.id === 'roles');
  const fwd = allSections.find(s => s.id === 'forwarding');
  const dev = allSections.find(s => s.id === 'devices');
  const ca = allSections.find(s => s.id === 'ca-analysis') || allSections.find(s => s.id === 'ca');
  const ai = allSections.find(s => s.id === 'ai-apps');
  const features = allSections.find(s => s.id === 'features');

  if (roles && roles.available) {
    const admins = (roles.rows || []).filter(r => String(r.RoleName || '').toLowerCase().includes('global admin')).length;
    strip.appendChild(card(admins, 'Global Admins', admins > 4 ? 'risk' : admins > 2 ? 'warn' : 'ok', () => openSectionWithPreset('roles', 'global-admins')));
  }

  if (fwd && fwd.available) {
    const total = (fwd.rows || []).length;
    strip.appendChild(card(total, 'Forwarding rules', total > 0 ? 'risk' : 'ok', () => openSectionWithPreset('forwarding', 'has-forwarding')));
  }

  if (dev && dev.available) {
    const stale = (dev.rows || []).filter(r => r.StaleDevice === true || toStr(r.StaleDevice).toLowerCase() === 'true').length;
    strip.appendChild(card(stale, 'Stale devices', stale > 0 ? 'warn' : 'ok', () => openSectionWithPreset('devices', 'stale-devices')));
  }

  if (ca && ca.available) {
    const rows = ca.rows || [];

    if (ca.id === 'ca-analysis') {
      const missing = rows.filter(r => toStr(r.Status).toLowerCase() === 'missing').length;
      strip.appendChild(card(missing + '/10', 'CA checks missing', missing > 5 ? 'risk' : missing > 2 ? 'warn' : missing > 0 ? 'warn' : 'ok', () => openSectionWithPreset('ca-analysis', 'ca-missing')));
    } else {
      const enabled = rows.filter(r => toStr(r.State).toLowerCase() === 'enabled').length;
      strip.appendChild(card(enabled, 'CA policies enabled', 'info', () => openSectionWithPreset('ca', '')));
    }
  }

  if (ai && ai.available) {
    const detected = (ai.rows || []).filter(r => toStr(r.Status).toLowerCase() === 'detected').length;
    strip.appendChild(card(detected, 'AI apps detected', detected > 0 ? 'warn' : 'ok', () => openSectionWithPreset('ai-apps', 'ai-detected')));
  }

  if (features && features.available) {
    const availableCount = (features.rows || []).filter(r => toStr(r.Available).toLowerCase() === 'yes').length;
    strip.appendChild(card(availableCount, 'Features available', 'info', () => openSectionWithPreset('features', 'features-available')));
  }

  const available = allSections.filter(s => s.available).length;
  strip.appendChild(card(available + '/' + allSections.length, 'Sections collected', 'neu', () => openSectionWithPreset('', '')));
}

function buildPanels() {
  const page = document.getElementById('mainPage');

  allSections.forEach(function(sec) {
    const panel = document.createElement('div');
    panel.className = 'section-panel';
    panel.id = 'panel-' + sec.id;

    if (!sec.available) {
      const naNote = esc(String(sec.note || 'Data not collected for this section — connection not available.'));
      const naSmall = esc(String(sec.note || ''));
      panel.innerHTML =
        '<div class="sec-header"><div class="sec-title-block"><h2>' + esc(sec.title) + '</h2><div class="sec-desc">' + esc(sec.description) + '</div></div></div>' +
        '<div class="callout callout-na">' + naNote + '</div>' +
        '<div class="unavail"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4m0 4h.01"/></svg><p>Section unavailable</p><small>' + naSmall + '</small></div>';
      page.appendChild(panel);
      return;
    }

    // header
    const header = document.createElement('div');
    header.className = 'sec-header';

    const titleBlock = document.createElement('div');
    titleBlock.className = 'sec-title-block';
    titleBlock.innerHTML = '<h2>' + esc(sec.title) + '</h2><div class="sec-desc">' + esc(sec.description) + '</div>';
    header.appendChild(titleBlock);

    if (sec.note) {
      const noteText = String(sec.note);
      const callout = document.createElement('div');
      callout.className = 'callout callout-' + (
        noteText.toLowerCase().match(/risk|warning|critical|caution|exposed/) ? 'risk' :
        noteText.toLowerCase().match(/ok|clean|no issue|none found/) ? 'ok' : 'warn'
      );
      callout.textContent = noteText;
      panel.appendChild(header);
      panel.appendChild(callout);
    }
    else {
      panel.appendChild(header);
    }

    // stat cards + table (stats need table controller for click-to-filter)
    let tblObj = null;
    if (sec.rows && sec.rows.length > 0) {
      tblObj = SectionTable(sec.id, sec.rows);
      sectionControllers[sec.id] = tblObj;
    }

    panel.appendChild(sectionStats(sec, tblObj));

    if (tblObj) {
      const actDiv = document.createElement('div');
      actDiv.className = 'sec-actions';
      actDiv.appendChild(tblObj.exportBtn);
      titleBlock.parentElement.appendChild(actDiv);
      panel.appendChild(tblObj.el);
    }
    else {
      const empty = document.createElement('div');
      empty.className='unavail';
      empty.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><rect x="3" y="3" width="18" height="18" rx="2.5"/><path d="M8 9h8M8 12h8M8 15h5"/></svg><p>No records found</p>';
      panel.appendChild(empty);
    }

    page.appendChild(panel);
  });
}

function activateTab(id) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === id));
  document.querySelectorAll('.section-panel').forEach(p => p.classList.toggle('active', p.id === 'panel-' + id));
}

document.querySelectorAll('.tab').forEach(function(tab){
  tab.addEventListener('click', function(){ activateTab(tab.dataset.tab); });
});

setupSettingsModal();
buildSummary();
buildPanels();

const now = new Date().toLocaleString();
document.getElementById('tbTime').textContent = 'Generated ' + now;
document.getElementById('ftrTime').textContent = 'Generated ' + now;

// activate first available section
const firstSec = allSections.find(s => s.available) || allSections[0];
if (firstSec) activateTab(firstSec.id);
</script>
</body>
</html>
"@

    $html = $html.Replace('__SECTIONS_JSON__', $jsSectionsForHtml).Replace('__UI_CONFIG_JSON__', $uiConfigJsonForHtml)

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
