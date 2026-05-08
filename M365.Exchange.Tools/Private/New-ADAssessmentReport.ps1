function New-ADAssessmentReport {
    <#
    .SYNOPSIS
        Generates a self-contained, multi-section HTML assessment report for Active Directory.
    .PARAMETER Sections
        Array of hashtables: {Id, Title, Description, Note, Available, Rows}
    .PARAMETER ReportTitle
        Title displayed in the report header.
    .PARAMETER PassThru
        If set, returns the report file path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Sections,

        [Parameter()]
        [string]$ReportTitle = 'Active Directory Assessment',

        [Parameter()]
        [switch]$PassThru
    )

    $settings        = Get-M365UiSettings
    $companyName     = [string]$settings.CompanyName
    $logoPath        = [string]$settings.LogoPath
    $brandingEnabled = [bool]$settings.HtmlBrandingEnabled
    $showName        = [bool]$settings.HtmlShowCompanyName
    $showLogo        = [bool]$settings.HtmlShowCompanyLogo
    $savePath        = if ([string]::IsNullOrWhiteSpace([string]$settings.ReportSavePath)) { $env:TEMP } else { [string]$settings.ReportSavePath }
    $fontFamily      = ([string]$settings.ReportFontFamily -replace '[^A-Za-z0-9,\- ]','').Trim()
    if ([string]::IsNullOrWhiteSpace($fontFamily)) { $fontFamily = 'Segoe UI' }

    $primary   = if (([string]$settings.ThemePrimaryColor)   -match '^#?[0-9A-Fa-f]{6}$') { if (([string]$settings.ThemePrimaryColor).StartsWith('#'))   { [string]$settings.ThemePrimaryColor }   else { "#$([string]$settings.ThemePrimaryColor)" }   } else { '#1e40af' }
    $secondary = if (([string]$settings.ThemeSecondaryColor) -match '^#?[0-9A-Fa-f]{6}$') { if (([string]$settings.ThemeSecondaryColor).StartsWith('#')) { [string]$settings.ThemeSecondaryColor } else { "#$([string]$settings.ThemeSecondaryColor)" } } else { '#0f172a' }

    $logoUri = ''
    if ($brandingEnabled -and $showLogo -and -not [string]::IsNullOrWhiteSpace($logoPath) -and (Test-Path -Path $logoPath)) {
        try { $logoUri = ([System.Uri](Resolve-Path -Path $logoPath -ErrorAction Stop).Path).AbsoluteUri } catch {}
    }
    $safeCompany = if ($brandingEnabled -and $showName -and -not [string]::IsNullOrWhiteSpace($companyName)) { [System.Net.WebUtility]::HtmlEncode($companyName) } else { '' }

    $brandHtml = ''
    if ($logoUri)     { $brandHtml += '<img class="brand-logo" src="' + [System.Net.WebUtility]::HtmlEncode($logoUri) + '" alt="logo">' }
    if ($safeCompany) { $brandHtml += '<span class="brand-name">' + $safeCompany + '</span>' }
    if ($brandHtml)   { $brandHtml += '<div class="tb-div"></div>' }

    $timestamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
    $titleToken = ($ReportTitle -replace '[^A-Za-z0-9\-_ ]','' -replace ' +','-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($titleToken)) { $titleToken = 'ADAssessment' }
    $companyToken = if (-not [string]::IsNullOrWhiteSpace($companyName)) { ($companyName -replace '[^A-Za-z0-9\-_ ]','' -replace ' +','-').Trim('-') } else { '' }
    $stem = if ($companyToken) { "$companyToken-$titleToken-$timestamp" } else { "$titleToken-$timestamp" }

    if (-not (Test-Path -Path $savePath)) { New-Item -Path $savePath -ItemType Directory -Force | Out-Null }
    $reportPath = Join-Path -Path $savePath -ChildPath "$stem.html"

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

    $safeTitle  = [System.Net.WebUtility]::HtmlEncode($ReportTitle)
    $genDate    = Get-Date -Format 'dd MMM yyyy HH:mm'
    $tabNavHtml = ($Sections | ForEach-Object {
        $id = [string]$_.Id
        $t  = [System.Net.WebUtility]::HtmlEncode([string]$_.Title)
        '<button class="tab" data-tab="' + $id + '" type="button">' + $t + '</button>'
    }) -join ''

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
  --red:#dc2626;--rbg:#fee2e2;
  --amber:#b45309;--abg:#fef3c7;
  --blue:#2563eb;--bbg:#dbeafe;
  --orange:#c2410c;--obg:#ffedd5;
  --purple:#7c3aed;--pbg:#ede9fe;
  --r:8px;
  --sh:0 1px 3px rgba(0,0,0,.08),0 1px 2px rgba(0,0,0,.05);
  --sh2:0 4px 6px -1px rgba(0,0,0,.07),0 2px 4px -2px rgba(0,0,0,.05);
}
html,body{height:100%;font-family:'$fontFamily',system-ui,-apple-system,sans-serif;background:var(--surf2);color:var(--t1);font-size:13px;line-height:1.5}
body{display:flex;flex-direction:column;min-height:100vh}
.topbar{background:linear-gradient(135deg,var(--pr) 0%,var(--sec) 100%);padding:0 20px;height:52px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 10px rgba(0,0,0,.28);flex-shrink:0;gap:12px}
.tb-left{display:flex;align-items:center;gap:10px;overflow:hidden;flex:1;min-width:0}
.brand-logo{height:28px;width:auto;max-width:130px;object-fit:contain;border-radius:3px;flex-shrink:0}
.brand-name{color:#e2e8f0;font-size:13px;font-weight:700;white-space:nowrap;flex-shrink:0}
.tb-div{width:1px;height:20px;background:rgba(255,255,255,.2);flex-shrink:0}
.tb-title{color:#fff;font-size:15px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tb-right{color:rgba(255,255,255,.65);font-size:11px;white-space:nowrap;flex-shrink:0}
.summary-strip{background:var(--surf);border-bottom:1px solid var(--bd);padding:10px 20px;display:flex;gap:10px;flex-wrap:wrap;flex-shrink:0}
.stat-card{display:flex;flex-direction:column;align-items:flex-start;padding:8px 14px;border-radius:var(--r);border:1px solid var(--bd);min-width:110px;background:var(--surf2)}
.stat-card.risk{border-color:#fca5a5;background:#fef2f2}.stat-card.risk .sv{color:var(--red)}
.stat-card.warn{border-color:#fcd34d;background:#fffbeb}.stat-card.warn .sv{color:var(--amber)}
.stat-card.ok{border-color:#86efac;background:#f0fdf4}.stat-card.ok .sv{color:var(--green)}
.stat-card.info{border-color:#93c5fd;background:#eff6ff}.stat-card.info .sv{color:var(--blue)}
.stat-card.neu{background:var(--surf2)}
.sv{font-size:22px;font-weight:800;line-height:1;margin-bottom:2px}
.sl{font-size:11px;color:var(--t3);font-weight:500}
.layout{display:flex;flex:1;overflow:hidden;min-height:0}
.sidebar{width:220px;background:var(--surf);border-right:1px solid var(--bd);overflow-y:auto;flex-shrink:0;padding:12px 0}
.sidebar-head{padding:6px 16px 4px;font-size:10px;font-weight:700;color:var(--t4);text-transform:uppercase;letter-spacing:.06em}
.tab{width:100%;background:none;border:none;text-align:left;padding:8px 16px;cursor:pointer;color:var(--t2);font-size:12px;font-family:inherit;border-left:3px solid transparent;transition:all .15s;line-height:1.3;word-break:break-word}
.tab:hover{background:var(--surf3);color:var(--t1)}
.tab.active{border-left-color:var(--pr);background:var(--surf3);color:var(--pr);font-weight:600}
.main{flex:1;overflow-y:auto;padding:20px;min-width:0}
.panel{display:none}.panel.active{display:block}
.panel-header{margin-bottom:14px}
.panel-title{font-size:17px;font-weight:700;color:var(--t1);margin-bottom:4px}
.panel-desc{color:var(--t3);font-size:12px;margin-bottom:8px}
.note-box{background:#fefce8;border:1px solid #fde68a;border-radius:var(--r);padding:8px 12px;color:#92400e;font-size:12px;margin-bottom:10px}
.sec-stats{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}
.mini-card{display:flex;flex-direction:column;align-items:flex-start;padding:6px 12px;border-radius:6px;border:1px solid var(--bd);min-width:90px;background:var(--surf2)}
.mini-card.risk{border-color:#fca5a5;background:#fef2f2}.mini-card.risk .mv{color:var(--red)}
.mini-card.warn{border-color:#fcd34d;background:#fffbeb}.mini-card.warn .mv{color:var(--amber)}
.mini-card.ok{border-color:#86efac;background:#f0fdf4}.mini-card.ok .mv{color:var(--green)}
.mini-card.info{border-color:#93c5fd;background:#eff6ff}.mini-card.info .mv{color:var(--blue)}
.mv{font-size:18px;font-weight:800;line-height:1;margin-bottom:1px}
.ml{font-size:10px;color:var(--t3);font-weight:500}
.toolbar{display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center}
.search-box{flex:1;min-width:160px;max-width:320px;padding:5px 10px;border:1px solid var(--bd);border-radius:6px;font-size:12px;font-family:inherit;outline:none}
.search-box:focus{border-color:var(--pr);box-shadow:0 0 0 2px rgba(30,64,175,.15)}
.btn{padding:5px 11px;border:1px solid var(--bd);border-radius:6px;background:var(--surf);cursor:pointer;font-size:11px;font-family:inherit;color:var(--t2);transition:background .12s}
.btn:hover{background:var(--surf3)}
.tbl-wrap{overflow-x:auto;border:1px solid var(--bd);border-radius:var(--r);background:var(--surf)}
table{width:100%;border-collapse:collapse;min-width:600px}
th{background:var(--surf3);font-size:11px;font-weight:700;color:var(--t2);padding:7px 10px;text-align:left;border-bottom:2px solid var(--bd2);white-space:nowrap;cursor:pointer;user-select:none;position:sticky;top:0;z-index:1}
th:hover{background:var(--bd)}
th .sort-icon{margin-left:4px;color:var(--t4);font-size:10px}
td{padding:6px 10px;border-bottom:1px solid var(--bd);color:var(--t2);vertical-align:top;font-size:12px;word-break:break-word;max-width:320px}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafc}
.chip{display:inline-flex;align-items:center;padding:2px 7px;border-radius:20px;font-size:10.5px;font-weight:600;white-space:nowrap}
.cg{background:var(--gbg);color:var(--green)}
.cr{background:var(--rbg);color:var(--red)}
.ca{background:var(--abg);color:var(--amber)}
.cb{background:var(--bbg);color:var(--blue)}
.co{background:var(--obg);color:var(--orange)}
.cp{background:var(--pbg);color:var(--purple)}
.cx{background:var(--surf3);color:var(--t3)}
.empty-state{text-align:center;padding:40px 20px;color:var(--t4)}
.empty-state svg{margin-bottom:8px;opacity:.4}
.empty-title{font-size:14px;font-weight:600;margin-bottom:4px;color:var(--t3)}
.rec-cell{font-size:11px;color:var(--t3);font-style:italic;max-width:300px}
.risk-cell{font-size:11px;color:var(--t3);max-width:280px}
</style>
</head>
<body>
<div class="topbar">
  <div class="tb-left">
    $brandHtml
    <span class="tb-title">$safeTitle</span>
  </div>
  <div class="tb-right">Generated $genDate</div>
</div>
<div id="summary-strip" class="summary-strip"></div>
<div class="layout">
  <nav class="sidebar">
    <div class="sidebar-head">Sections</div>
    $tabNavHtml
  </nav>
  <div class="main" id="main-content"></div>
</div>
<script>
const allSections = $jsSections;

function toStr(v){
  if(v===null||v===undefined)return'';
  if(typeof v==='boolean')return v?'True':'False';
  return String(v);
}

function renderCell(raw,cl){
  const td=document.createElement('td');
  const s=toStr(raw);
  if(s===''){{td.textContent='';return td;}}

  // Status (Pass/OK/Warning/Critical/Info)
  if(cl==='status'||cl==='osstatus'){
    const chip=document.createElement('span');
    const rl=s.toLowerCase();
    chip.className='chip '+(rl==='pass'||rl==='ok'?'cg':rl==='warning'?'ca':rl==='critical'?'cr':rl==='info'?'cb':'cx');
    chip.textContent=s;td.appendChild(chip);return td;
  }

  // Severity (Info/Low/Medium/High/Critical)
  if(cl==='severity'){
    const chip=document.createElement('span');
    const rl=s.toLowerCase();
    chip.className='chip '+(rl==='critical'?'cr':rl==='high'?'co':rl==='medium'?'ca':rl==='low'?'cg':'cb');
    chip.textContent=s;td.appendChild(chip);return td;
  }

  // Boolean-like values
  if(s==='True'||s==='False'){
    const chip=document.createElement('span');
    chip.className='chip '+(s==='True'?'cg':'cx');
    chip.textContent=s;td.appendChild(chip);return td;
  }
  if(typeof raw==='boolean'){
    const chip=document.createElement('span');
    chip.className='chip '+(raw?'cg':'cx');
    chip.textContent=raw?'True':'False';td.appendChild(chip);return td;
  }

  // SYSVOLReplication chip
  if(cl==='sysvolreplication'){
    const chip=document.createElement('span');
    chip.className='chip '+(s.includes('DFSR')?'cg':'cr');
    chip.textContent=s;td.appendChild(chip);return td;
  }

  // Pingable chip
  if(cl==='pingable'){
    const chip=document.createElement('span');
    chip.className='chip '+(s==='True'?'cg':'cr');
    chip.textContent=s==='True'?'Reachable':'Unreachable';td.appendChild(chip);return td;
  }

  // Type column
  if(cl==='type'){
    const chip=document.createElement('span');
    chip.className='chip cb';chip.textContent=s;td.appendChild(chip);return td;
  }

  // Dates
  if(s.match(/^\d{4}-\d{2}-\d{2}T/)){
    const d=new Date(s);
    td.textContent=isNaN(d)?s:d.toLocaleDateString();return td;
  }

  // Risk / Recommendation columns — italicised small text
  if(cl==='risk'){td.className='risk-cell';td.textContent=s;return td;}
  if(cl==='recommendation'){td.className='rec-cell';td.textContent=s;return td;}

  td.textContent=s;
  return td;
}

function sectionStats(sec){
  const row=document.createElement('div');
  row.className='sec-stats';
  const rows=sec.rows||[];
  function card(v,label,cls){
    const c=document.createElement('div');c.className='mini-card '+(cls||'');
    const sv=document.createElement('div');sv.className='mv';sv.textContent=v;
    const sl=document.createElement('div');sl.className='ml';sl.textContent=label;
    c.appendChild(sv);c.appendChild(sl);return c;
  }
  if(!sec.available||rows.length===0){row.appendChild(card('—','No data',''));return row;}

  const id=sec.id;

  if(id==='domain-summary'){
    const crit=rows.filter(r=>toStr(r.Status).toLowerCase()==='critical').length;
    const warn=rows.filter(r=>toStr(r.Status).toLowerCase()==='warning').length;
    const ok  =rows.filter(r=>toStr(r.Status).toLowerCase()==='ok'||toStr(r.Status).toLowerCase()==='pass').length;
    row.appendChild(card(crit,'Critical issues',crit>0?'risk':'ok'));
    row.appendChild(card(warn,'Warnings',warn>0?'warn':''));
    row.appendChild(card(ok,'Checks OK','ok'));
  } else if(id==='dc-inventory'){
    const total =rows.length;
    const eol   =rows.filter(r=>toStr(r.OSStatus).toLowerCase()==='critical').length;
    const warn  =rows.filter(r=>toStr(r.OSStatus).toLowerCase()==='warning').length;
    const rodcs =rows.filter(r=>toStr(r.IsReadOnlyDC).toLowerCase()==='true').length;
    row.appendChild(card(total,'Domain controllers','info'));
    row.appendChild(card(eol,'EOL OS',eol>0?'risk':'ok'));
    row.appendChild(card(warn,'OS nearing EOL',warn>0?'warn':''));
    row.appendChild(card(rodcs,'Read-only DCs',''));
  } else if(id==='replication'){
    const failed=rows.filter(r=>toStr(r.Status).toLowerCase()==='critical').length;
    const warn  =rows.filter(r=>toStr(r.Status).toLowerCase()==='warning').length;
    const ok    =rows.filter(r=>toStr(r.Status).toLowerCase()==='ok').length;
    row.appendChild(card(failed,'Replication failures',failed>0?'risk':'ok'));
    row.appendChild(card(warn,'Warnings',warn>0?'warn':''));
    row.appendChild(card(ok,'Healthy links','ok'));
  } else if(id==='dns-health'){
    const issues=rows.filter(r=>toStr(r.Status).toLowerCase()==='warning'||toStr(r.Status).toLowerCase()==='critical').length;
    const zones =rows.filter(r=>toStr(r.ZoneType)!=='Global Setting').length;
    row.appendChild(card(zones,'DNS zones','info'));
    row.appendChild(card(issues,'Zones with issues',issues>0?'warn':'ok'));
  } else if(id==='sites-services'){
    const sites =rows.filter(r=>toStr(r.Type)==='Site').length;
    const links =rows.filter(r=>toStr(r.Type)==='Site Link').length;
    const issues=rows.filter(r=>toStr(r.Status).toLowerCase()==='warning').length;
    row.appendChild(card(sites,'Sites','info'));
    row.appendChild(card(links,'Site links','info'));
    row.appendChild(card(issues,'Issues',issues>0?'warn':'ok'));
  } else if(id==='security-posture'){
    const crit=rows.filter(r=>toStr(r.Status).toLowerCase()==='critical').length;
    const warn=rows.filter(r=>toStr(r.Status).toLowerCase()==='warning').length;
    const pass=rows.filter(r=>toStr(r.Status).toLowerCase()==='pass').length;
    row.appendChild(card(crit,'Critical',crit>0?'risk':'ok'));
    row.appendChild(card(warn,'Warnings',warn>0?'warn':''));
    row.appendChild(card(pass,'Passed','ok'));
  } else if(id==='operational-risk'){
    const crit=rows.filter(r=>toStr(r.Severity).toLowerCase()==='critical'||toStr(r.Severity).toLowerCase()==='high').length;
    const med =rows.filter(r=>toStr(r.Severity).toLowerCase()==='medium').length;
    const info=rows.filter(r=>toStr(r.Severity).toLowerCase()==='info'||toStr(r.Severity).toLowerCase()==='low').length;
    row.appendChild(card(crit,'Critical/High',crit>0?'risk':'ok'));
    row.appendChild(card(med,'Medium',med>0?'warn':''));
    row.appendChild(card(info,'Low/Info','info'));
  } else {
    row.appendChild(card(rows.length,'Total records','info'));
  }
  return row;
}

function buildSummary(){
  const strip=document.getElementById('summary-strip');
  if(!strip)return;
  function card(v,label,cls){
    const c=document.createElement('div');c.className='stat-card '+(cls||'neu');
    const sv=document.createElement('div');sv.className='sv';sv.textContent=v;
    const sl=document.createElement('div');sl.className='sl';sl.textContent=label;
    c.appendChild(sv);c.appendChild(sl);return c;
  }
  const sec=id=>allSections.find(s=>s.id===id);
  const rows=id=>(sec(id)&&sec(id).available)?sec(id).rows||[]:[];

  // DC count
  const dcRows=rows('dc-inventory');
  const dcCount=dcRows.length;
  const eolDCs=dcRows.filter(r=>toStr(r.OSStatus).toLowerCase()==='critical').length;
  strip.appendChild(card(dcCount,'Domain controllers','info'));
  if(eolDCs>0) strip.appendChild(card(eolDCs,'DCs on EOL OS','risk'));

  // Replication
  const repRows=rows('replication');
  const repFail=repRows.filter(r=>toStr(r.Status).toLowerCase()==='critical').length;
  if(repRows.length>0) strip.appendChild(card(repFail,'Replication failures',repFail>0?'risk':'ok'));

  // Security posture
  const secRows=rows('security-posture');
  const secCrit=secRows.filter(r=>toStr(r.Status).toLowerCase()==='critical').length;
  const secWarn=secRows.filter(r=>toStr(r.Status).toLowerCase()==='warning').length;
  if(secRows.length>0){
    strip.appendChild(card(secCrit,'Security critical',secCrit>0?'risk':'ok'));
    strip.appendChild(card(secWarn,'Security warnings',secWarn>0?'warn':''));
  }

  // Operational risk
  const opRows=rows('operational-risk');
  const opHigh=opRows.filter(r=>toStr(r.Severity).toLowerCase()==='critical'||toStr(r.Severity).toLowerCase()==='high').length;
  if(opRows.length>0) strip.appendChild(card(opHigh,'High/Critical risks',opHigh>0?'risk':'ok'));
}

function buildTable(rows,cols){
  const wrap=document.createElement('div');wrap.className='tbl-wrap';
  if(!rows||rows.length===0){
    wrap.innerHTML='<div class="empty-state"><div class="empty-title">No records</div></div>';
    return wrap;
  }
  const tbl=document.createElement('table');
  // Head
  const thead=document.createElement('thead');
  const htr=document.createElement('tr');
  cols.forEach((c,ci)=>{
    const th=document.createElement('th');
    th.innerHTML=c+'<span class="sort-icon">&#8597;</span>';
    th.dataset.col=ci;
    th.addEventListener('click',()=>sortTable(tbl,ci,th));
    htr.appendChild(th);
  });
  thead.appendChild(htr);tbl.appendChild(thead);
  // Body
  const tbody=document.createElement('tbody');
  rows.forEach(r=>{
    const tr=document.createElement('tr');
    cols.forEach(c=>{
      const cl=c.toLowerCase().replace(/[^a-z0-9]/g,'');
      const raw=r[c]!==undefined?r[c]:(r[Object.keys(r).find(k=>k.toLowerCase().replace(/[^a-z0-9]/g,'')===cl)]??'');
      tr.appendChild(renderCell(raw,cl));
    });
    tbody.appendChild(tr);
  });
  tbl.appendChild(tbody);wrap.appendChild(tbl);
  return wrap;
}

function sortTable(tbl,ci,th){
  const tbody=tbl.querySelector('tbody');
  const rows=Array.from(tbody.querySelectorAll('tr'));
  const asc=th.dataset.asc!=='true';
  th.dataset.asc=asc;
  tbl.querySelectorAll('th').forEach(h=>h.querySelector('.sort-icon').textContent='\u2195');
  th.querySelector('.sort-icon').textContent=asc?'\u2191':'\u2193';
  rows.sort((a,b)=>{
    const av=a.cells[ci]?.textContent||'';
    const bv=b.cells[ci]?.textContent||'';
    const an=parseFloat(av),bn=parseFloat(bv);
    if(!isNaN(an)&&!isNaN(bn))return asc?an-bn:bn-an;
    return asc?av.localeCompare(bv):bv.localeCompare(av);
  });
  rows.forEach(r=>tbody.appendChild(r));
}

function filterTable(tbl,q){
  const lq=q.toLowerCase();
  Array.from(tbl.querySelectorAll('tbody tr')).forEach(r=>{
    r.style.display=r.textContent.toLowerCase().includes(lq)?'':'none';
  });
}

function exportCSV(rows,cols,title){
  const esc=v=>'"'+String(v??'').replace(/"/g,'""')+'"';
  const lines=[cols.map(esc).join(',')];
  rows.forEach(r=>lines.push(cols.map(c=>{
    const cl=c.toLowerCase().replace(/[^a-z0-9]/g,'');
    const raw=r[c]!==undefined?r[c]:(r[Object.keys(r).find(k=>k.toLowerCase().replace(/[^a-z0-9]/g,'')===cl)]??'');
    return esc(raw);
  }).join(',')));
  const blob=new Blob([lines.join('\r\n')],{type:'text/csv'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download=(title.replace(/[^A-Za-z0-9\-_ ]/g,'')||'export')+'.csv';
  a.click();URL.revokeObjectURL(a.href);
}

function buildSection(sec){
  const panel=document.createElement('div');
  panel.id='panel-'+sec.id;panel.className='panel';

  const hdr=document.createElement('div');hdr.className='panel-header';
  const t=document.createElement('div');t.className='panel-title';t.textContent=sec.title;
  const d=document.createElement('div');d.className='panel-desc';d.textContent=sec.description;
  hdr.appendChild(t);hdr.appendChild(d);

  if(sec.note){
    const nb=document.createElement('div');nb.className='note-box';nb.textContent=sec.note;
    hdr.appendChild(nb);
  }
  panel.appendChild(hdr);

  if(!sec.available){
    const es=document.createElement('div');es.className='empty-state';
    es.innerHTML='<div class="empty-title">Data not available</div><div>'+sec.note+'</div>';
    panel.appendChild(es);return panel;
  }

  const rows=sec.rows||[];
  panel.appendChild(sectionStats(sec));

  if(rows.length===0){
    const es=document.createElement('div');es.className='empty-state';
    es.innerHTML='<div class="empty-title">No records found</div>';
    panel.appendChild(es);return panel;
  }

  const cols=Object.keys(rows[0]);

  // Toolbar
  const tb=document.createElement('div');tb.className='toolbar';
  const si=document.createElement('input');
  si.type='text';si.placeholder='Filter...';si.className='search-box';
  tb.appendChild(si);

  const expBtn=document.createElement('button');expBtn.className='btn';expBtn.textContent='Export CSV';
  tb.appendChild(expBtn);
  panel.appendChild(tb);

  const tblWrap=buildTable(rows,cols);
  panel.appendChild(tblWrap);

  si.addEventListener('input',()=>{
    const tbl=tblWrap.querySelector('table');
    if(tbl) filterTable(tbl,si.value);
  });
  expBtn.addEventListener('click',()=>exportCSV(rows,cols,sec.title));

  return panel;
}

function init(){
  buildSummary();
  const main=document.getElementById('main-content');
  allSections.forEach(sec=>main.appendChild(buildSection(sec)));
  document.querySelectorAll('.tab').forEach(tab=>{
    tab.addEventListener('click',()=>{
      document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
      document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
      tab.classList.add('active');
      const panel=document.getElementById('panel-'+tab.dataset.tab);
      if(panel) panel.classList.add('active');
    });
  });
  // Activate first tab
  const first=document.querySelector('.tab');
  if(first) first.click();
}

document.addEventListener('DOMContentLoaded',init);
</script>
</body>
</html>
"@

    Set-Content -Path $reportPath -Value $html -Encoding UTF8
    Write-Host "AD Assessment report saved: $reportPath" -ForegroundColor Green

    if ($PassThru) { return $reportPath }
    return $reportPath
}
