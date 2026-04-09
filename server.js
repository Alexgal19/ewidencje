'use strict';

const express = require('express');
const multer  = require('multer');
const path    = require('path');

const { parseGps, aggregate, aggregateActual, getPreviousWorkingDay } = require('./src/gpsParser');
const { generateExcel } = require('./src/excelGenerator');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.use(express.static(path.join(__dirname, 'public')));

// ── HTML page (verbatim from app.py) ─────────────────────────────────────────

/* eslint-disable no-useless-escape */
const HTML_PAGE = String.raw`<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Ewidencja Przebiegu Pojazdu &#x2013; Smart Work</title>
<link rel="manifest" href="/manifest.json">
<meta name="theme-color" content="#10B981">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Ewidencja">
<link rel="apple-touch-icon" href="/icons/icon-192.png">
<style>
:root{--bg:#0B1120;--surface:#111827;--card:#1C2535;--border:#2D3748;--accent:#10B981;--accent2:#059669;--text:#F1F5F9;--muted:#94A3B8;--gray:#6B7280;--err:#EF4444;--radius:14px}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:32px 16px 60px}
body::before{content:'';position:fixed;inset:0;background:radial-gradient(ellipse 700px 500px at 20% 10%,rgba(16,185,129,.07) 0%,transparent 60%),radial-gradient(ellipse 600px 400px at 80% 80%,rgba(16,185,129,.05) 0%,transparent 60%);pointer-events:none;z-index:0}
.wrap{width:100%;max-width:680px;position:relative;z-index:1;animation:up .5s ease}
@keyframes up{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:translateY(0)}}
.header{margin-bottom:32px}
.logo{display:flex;align-items:center;gap:14px;margin-bottom:14px}
.logo-icon{width:52px;height:52px;border-radius:14px;background:linear-gradient(135deg,var(--accent),#065F46);display:flex;align-items:center;justify-content:center;font-size:26px;box-shadow:0 0 28px rgba(16,185,129,.35);flex-shrink:0}
.logo-text{font-size:11px;letter-spacing:.2em;text-transform:uppercase;color:var(--accent);font-weight:600}
h1{font-size:26px;font-weight:800;line-height:1.25}
.subtitle{font-size:13px;color:var(--muted);margin-top:6px;line-height:1.6}
.card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);margin-bottom:16px;overflow:hidden;transition:border-color .2s}
.card:focus-within{border-color:rgba(16,185,129,.5)}
.card-head{padding:16px 22px 12px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.card-head .icon{width:32px;height:32px;border-radius:8px;background:rgba(16,185,129,.12);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}
.card-head-text h3{font-size:13px;font-weight:600}
.card-head-text p{font-size:11px;color:var(--muted);margin-top:1px}
.card-body{padding:18px 22px 20px}
.upload-zone{border:2px dashed var(--border);border-radius:10px;padding:28px 20px;cursor:pointer;text-align:center;transition:all .22s;position:relative;background:var(--surface)}
.upload-zone input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.upload-zone:hover,.upload-zone.drag{border-color:var(--accent);background:rgba(16,185,129,.04)}
.upload-zone.has-file{border-color:var(--accent);border-style:solid;background:rgba(16,185,129,.06)}
.upload-icon{font-size:32px;margin-bottom:8px;display:block}
.upload-title{font-size:14px;font-weight:600}
.upload-hint{font-size:11px;color:var(--muted);margin-top:4px;font-family:Consolas,monospace}
.upload-badge{display:none;margin-top:10px;background:rgba(16,185,129,.15);color:var(--accent);border:1px solid rgba(16,185,129,.3);border-radius:20px;padding:3px 10px;font-size:11px;font-family:Consolas,monospace}
.has-file .upload-badge{display:inline-block}
.has-file .upload-hint{color:var(--accent);opacity:.8}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
@media(max-width:500px){.grid2{grid-template-columns:1fr}}
.field{display:flex;flex-direction:column;gap:6px}
.field label{font-size:11px;font-weight:600;color:var(--muted);letter-spacing:.05em}
.field input{background:var(--surface);border:1.5px solid var(--border);border-radius:8px;padding:11px 14px;color:var(--text);font-family:inherit;font-size:13px;outline:none;transition:border-color .18s,box-shadow .18s}
.field input:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(16,185,129,.12)}
.field input::placeholder{color:var(--muted)}
.refuel-wrap{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px;min-height:32px}
.tag{display:flex;align-items:center;gap:6px;background:rgba(245,158,11,.12);border:1px solid rgba(245,158,11,.35);color:#FCD34D;border-radius:20px;padding:4px 10px 4px 12px;font-size:12px;animation:tagIn .2s ease}
@keyframes tagIn{from{opacity:0;transform:scale(.8)}to{opacity:1;transform:scale(1)}}
.tag .rm{cursor:pointer;opacity:.7;font-size:14px}
.tag .rm:hover{opacity:1;color:var(--err)}
.refuel-hint{font-size:11px;color:var(--muted);margin-top:6px}
.btn{width:100%;padding:16px;background:linear-gradient(135deg,var(--accent),var(--accent2));color:#052e16;font-size:15px;font-weight:800;letter-spacing:.03em;border:none;border-radius:var(--radius);cursor:pointer;transition:all .22s;display:flex;align-items:center;justify-content:center;gap:10px;position:relative;overflow:hidden}
.btn:hover{transform:translateY(-1px);box-shadow:0 10px 30px rgba(16,185,129,.3)}
.btn:active{transform:translateY(0)}
.btn:disabled{opacity:.5;cursor:not-allowed;transform:none;box-shadow:none}
.spinner{width:18px;height:18px;border:2.5px solid rgba(5,46,22,.3);border-top-color:#052e16;border-radius:50%;animation:spin .7s linear infinite;display:none}
@keyframes spin{to{transform:rotate(360deg)}}
.progress{height:4px;background:var(--border);border-radius:2px;margin-top:14px;overflow:hidden;display:none}
.progress-bar{height:100%;background:linear-gradient(90deg,var(--accent),#34D399);border-radius:2px;width:0%;transition:width .4s}
.progress.show{display:block}
.status{display:none;margin-top:14px;border-radius:10px;padding:14px 16px;font-size:13px;animation:up .3s ease;align-items:flex-start;gap:10px}
.status.ok{display:flex;background:rgba(16,185,129,.08);border:1px solid rgba(16,185,129,.3);color:var(--accent)}
.status.err{display:flex;background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.3);color:#FCA5A5}
.status .si{font-size:18px;flex-shrink:0}
.info-box{display:flex;gap:10px;align-items:flex-start;background:rgba(255,255,255,.03);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-top:16px;font-size:12px;color:var(--muted);line-height:1.6}
.install-btn{display:none;width:100%;padding:13px;background:transparent;color:var(--accent);font-size:13px;font-weight:700;letter-spacing:.03em;border:1.5px solid rgba(16,185,129,.4);border-radius:var(--radius);cursor:pointer;transition:all .22s;align-items:center;justify-content:center;gap:8px;margin-bottom:12px}
.install-btn:hover{background:rgba(16,185,129,.08);border-color:var(--accent)}
.install-btn.visible{display:flex}
</style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <div class="logo">
      <div class="logo-icon">&#x1F697;</div>
      <div>
        <div class="logo-text">Smart Work Sp. z o.o.</div>
      </div>
    </div>
    <h1>Generator Ewidencji<br>Przebiegu Pojazdu</h1>
    <p class="subtitle">Wgraj raport GPS &#x2192; uzupe&#x142;nij dane &#x2192; pobierz gotowy plik Excel.<br>Weekendy automatycznie sumowane na pi&#x105;tek. Dzia&#x142;a bez LibreOffice.</p>
  </div>

  <div class="card" id="gps-card">
    <div class="card-head">
      <div class="icon">&#x1F4CD;</div>
      <div class="card-head-text">
        <h3>Plik GPS</h3>
        <p>Raport z trackera pojazdu (.XLS lub .XLSX)</p>
      </div>
    </div>
    <div class="card-body">
      <div class="upload-zone" id="zone">
        <input type="file" id="gps_file" accept=".xls,.xlsx" onchange="handleFile(this)">
        <span class="upload-icon">&#x1F4C2;</span>
        <div class="upload-title" id="file-title">Kliknij lub przeci&#x105;gnij plik tutaj</div>
        <div class="upload-hint" id="file-hint">.XLS &#xB7; .XLSX &#xB7; max 50 MB</div>
        <div class="upload-badge" id="file-badge">&#x2713; plik wgrany</div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-head">
      <div class="icon">&#x1F464;</div>
      <div class="card-head-text">
        <h3>Dane kierowcy</h3>
        <p>Imi&#x119; i nazwisko pojawi si&#x119; we wszystkich wierszach z km</p>
      </div>
    </div>
    <div class="card-body">
      <div class="grid2">
        <div class="field">
          <label>NAZWISKO I IMI&#x118;</label>
          <input type="text" id="driver_name" placeholder="np. Kowalski Jan" autocomplete="off">
        </div>
        <div class="field">
          <label>STAN LICZNIKA NA POCZ&#x104;TKU MIESI&#x104;CA (KM)</label>
          <input type="number" id="odometer_start" placeholder="np. 248349" min="0">
        </div>
      </div>
      <div class="field" style="margin-top:14px;">
        <label>CEL WYJAZDU</label>
        <input type="text" id="trip_purpose" placeholder="np. dow&#xF3;z/odbi&#xF3;r pracownik&#xF3;w" autocomplete="off">
      </div>
    </div>
  </div>

  <div class="card" id="target-korekta-card">
    <div class="card-head">
      <div class="icon">&#x1F3AF;</div>
      <div class="card-head-text">
        <h3>Korekta ko&#x144;cowego przebiegu (Opcjonalnie)</h3>
        <p>Wpisz r&#x119;cznie docelowy stan licznika, aby zr&#xF3;wna&#x107; trasy</p>
      </div>
    </div>
    <div class="card-body">
      <div class="grid2">
        <div class="field">
          <label>DOCELOWY STAN LICZNIKA KO&#x143;COWY</label>
          <input type="number" id="target_odometer" placeholder="np. 250000" min="0">
        </div>
        <div class="field" style="justify-content: flex-end; padding-bottom: 8px;">
          <label style="display:flex; align-items:center; gap:8px; cursor:pointer; font-size:13px; color:var(--text); font-weight:600;">
            <input type="checkbox" id="adjust_mileage" checked style="width:18px; height:18px; accent-color: var(--accent);">
            Podci&#x105;gnij trasy, aby zr&#xF3;wna&#x107;
          </label>
        </div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-head">
      <div class="icon">&#x26FD;</div>
      <div class="card-head-text">
        <h3>Dni tankowania</h3>
        <p>Wpisz daty &#x2013; pojawi&#x105; si&#x119; w kolumnie Uwagi jako "tankowanie"</p>
      </div>
    </div>
    <div class="card-body">
      <div class="field">
        <label>WPISZ DAT&#x118; I NACI&#x15A;NIJ ENTER LUB SPACJ&#x118;</label>
        <input type="text" id="refuel_input" placeholder="np. 5.12 lub 05.12.2025" autocomplete="off">
      </div>
      <div class="refuel-wrap" id="refuel-tags"></div>
      <div class="refuel-hint">&#x1F4CC; Format: DD.MM lub DD.MM.RRRR &nbsp;|&nbsp; Mo&#x17C;esz doda&#x107; wiele dat</div>
    </div>
  </div>

  <div class="card">
    <div class="card-head">
      <div class="icon">&#x1F4C5;</div>
      <div class="card-head-text">
        <h3>Tryb zapisu przebiegu</h3>
        <p>Wybierz jak maj&#x105; by&#x107; przypisane kilometry w ewidencji</p>
      </div>
    </div>
    <div class="card-body">
      <label style="display:flex;align-items:center;gap:10px;cursor:pointer;font-size:13px;color:var(--text);font-weight:600;">
        <input type="checkbox" id="use_actual_days" style="width:18px;height:18px;accent-color:var(--accent);">
        Zapisz przebieg w faktycznych dniach jazdy (bez przesuwania na pi&#x105;tek)
      </label>
      <p style="font-size:11px;color:var(--muted);margin-top:8px;line-height:1.6;">
        <strong style="color:var(--text);">Domy&#x15B;lnie (odznaczone):</strong> km z weekendu/&#x15B;wi&#x105;t &#x2192; poprzedni pi&#x105;tek roboczy.<br>
        <strong style="color:var(--text);">Po zaznaczeniu:</strong> km zapisane dok&#x142;adnie w dniu, w kt&#xF3;rym je&#x17A;dzi&#x142;o auto.
      </p>
    </div>
  </div>

  <button class="install-btn" id="install-btn" onclick="installApp()">&#x2B07; Zainstaluj aplikacj&#x119; na urz&#x105;dzeniu</button>
  <button class="btn" id="btn" onclick="generate()">

    <div class="spinner" id="spinner"></div>
    <span id="btn-text">&#x26A1; Generuj ewidencj&#x119; Excel</span>
  </button>
  <div class="progress" id="progress"><div class="progress-bar" id="pbar"></div></div>
  <div class="status" id="status"><div class="si" id="s-icon"></div><div id="s-msg"></div></div>

  <div class="info-box">
    <div style="font-size:16px;flex-shrink:0">&#x1F4A1;</div>
    <div>Aplikacja odczytuje dane GPS, sumuje kilometry per dzie&#x144; i tworzy plik Excel ze stron&#x105; tytu&#x142;ow&#x105; i arkuszem rozliczeniowym. Dane z soboty i niedzieli przenoszone s&#x105; automatycznie na poprzedni pi&#x105;tek. <strong style="color:#10B981">Nie wymaga LibreOffice.</strong></div>
  </div>
</div>
<script>
// ── PWA ──────────────────────────────────────────────────────────────────────
let _installPrompt=null;
if('serviceWorker' in navigator){
  window.addEventListener('load',()=>navigator.serviceWorker.register('/sw.js').catch(()=>{}));
}
window.addEventListener('beforeinstallprompt',e=>{
  e.preventDefault();_installPrompt=e;
  document.getElementById('install-btn').classList.add('visible');
});
window.addEventListener('appinstalled',()=>{
  document.getElementById('install-btn').classList.remove('visible');
  _installPrompt=null;
});
function installApp(){
  if(!_installPrompt)return;
  _installPrompt.prompt();
  _installPrompt.userChoice.then(()=>{_installPrompt=null;document.getElementById('install-btn').classList.remove('visible')});
}
// ── App ───────────────────────────────────────────────────────────────────────
const refuelDates=[];
const zone=document.getElementById('zone');
zone.addEventListener('dragover',e=>{e.preventDefault();zone.classList.add('drag')});
zone.addEventListener('dragleave',()=>zone.classList.remove('drag'));
zone.addEventListener('drop',e=>{e.preventDefault();zone.classList.remove('drag');if(e.dataTransfer.files.length){document.getElementById('gps_file').files=e.dataTransfer.files;handleFile(document.getElementById('gps_file'))}});
function handleFile(input){if(input.files&&input.files[0]){const name=input.files[0].name;zone.classList.add('has-file');document.getElementById('file-title').textContent=name;document.getElementById('file-hint').textContent=(input.files[0].size/1024/1024).toFixed(2)+' MB';document.getElementById('file-badge').textContent='\u2713 '+name;document.getElementById('gps-card').style.borderColor='rgba(16,185,129,.6)'}}
document.getElementById('refuel_input').addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '||e.key===','){e.preventDefault();const v=e.target.value.trim();if(v){addTag(v);e.target.value=''}}});
document.getElementById('refuel_input').addEventListener('blur',e=>{const v=e.target.value.trim();if(v){addTag(v);e.target.value=''}});
function addTag(val){val=val.replace(/[,;]/g,'').trim();if(!val||refuelDates.includes(val))return;if(!/^\d{1,2}[.\-\/]\d{1,2}([.\-\/]\d{2,4})?$/.test(val)){showStatus('err','\u26A0','Nieprawid\u0142owy format: '+val+'. U\u017Cyj DD.MM');return}refuelDates.push(val);const wrap=document.getElementById('refuel-tags');const tag=document.createElement('div');tag.className='tag';tag.innerHTML='&#x26FD; '+val+' <span class="rm" data-val="'+val+'">&#x2715;</span>';tag.querySelector('.rm').onclick=function(){refuelDates.splice(refuelDates.indexOf(this.dataset.val),1);this.parentElement.remove()};wrap.appendChild(tag)}
let progIv=null;
function startProgress(){const wrap=document.getElementById('progress');const bar=document.getElementById('pbar');wrap.classList.add('show');bar.style.width='0%';let w=0;progIv=setInterval(()=>{w=Math.min(w+Math.random()*7,88);bar.style.width=w+'%'},200)}
function stopProgress(){clearInterval(progIv);document.getElementById('pbar').style.width='100%';setTimeout(()=>document.getElementById('progress').classList.remove('show'),500)}
function showStatus(type,icon,msg){const el=document.getElementById('status');el.className='status '+type;document.getElementById('s-icon').textContent=icon;document.getElementById('s-msg').textContent=msg}
async function generate(){
  const file=document.getElementById('gps_file').files[0];
  if(!file){showStatus('err','\u26A0','Wybierz plik GPS.');return}
  const btn=document.getElementById('btn');const spin=document.getElementById('spinner');const btxt=document.getElementById('btn-text');
  btn.disabled=true;spin.style.display='block';btxt.textContent='Przetwarzanie...';
  document.getElementById('status').className='status';
  startProgress();
  try{
    const form=new FormData();
    form.append('gps_file',file);
    form.append('driver_name',document.getElementById('driver_name').value);
    form.append('odometer_start',document.getElementById('odometer_start').value||'0');
    form.append('refuel_dates',refuelDates.join(','));
    form.append('target_odometer',document.getElementById('target_odometer').value||'0');
    form.append('adjust_mileage',document.getElementById('adjust_mileage').checked);
    form.append('use_actual_days',document.getElementById('use_actual_days').checked);
    form.append('trip_purpose',document.getElementById('trip_purpose').value.trim());
    const resp=await fetch('/generate',{method:'POST',body:form});
    stopProgress();
    if(!resp.ok){const data=await resp.json().catch(()=>({}));throw new Error(data.error||'B\u0142\u0105d serwera ('+resp.status+')')}
    const blob=await resp.blob();
    const cd=resp.headers.get('Content-Disposition')||'';
    const m=cd.match(/filename[^=\n]*=(["']?)(.+?)\1/);
    const name=m?m[2]:'ewidencja.xlsx';
    const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
    showStatus('ok','\u2713','Plik "'+name+'" pobrany!');
  }catch(err){stopProgress();showStatus('err','\u2717','B\u0142\u0105d: '+err.message)}
  finally{btn.disabled=false;spin.style.display='none';btxt.textContent='\u26A1 Generuj ewidencj\u0119 Excel'}
}
</script>
</body>
</html>`;
/* eslint-enable no-useless-escape */

// ── Routes ────────────────────────────────────────────────────────────────────

app.get('/', (req, res) => {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.send(HTML_PAGE);
});

app.post('/generate', upload.single('gps_file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Brak pliku GPS' });
        }

        const buffer   = req.file.buffer;
        const filename = req.file.originalname || 'upload.xls';

        const driver       = (req.body.driver_name      || '').trim();
        const odoRaw       = (req.body.odometer_start   || '0').trim();
        const refuelRaw    = (req.body.refuel_dates      || '').trim();
        const targetOdoRaw = (req.body.target_odometer  || '0').trim();
        const adjustMil    = req.body.adjust_mileage === 'true';
        const useActual    = req.body.use_actual_days   === 'true';
        const tripPurpose  = (req.body.trip_purpose     || '').trim();

        const odometer       = parseInt(odoRaw.replace(/[\s,]/g, ''), 10) || 0;
        const targetOdometer = parseInt(targetOdoRaw.replace(/[\s,]/g, ''), 10) || 0;

        const { plate, carModel, dateFrom, dateTo, dayGroups } = await parseGps(buffer, filename);

        if (!dayGroups || !dayGroups.length) {
            return res.status(400).json({ error: 'Nie znaleziono danych GPS w pliku.' });
        }

        let agg = useActual ? aggregateActual(dayGroups) : aggregate(dayGroups);

        // Proportional km adjustment
        if (adjustMil && targetOdometer > odometer) {
            const expectedGpsKm = targetOdometer - odometer;
            const currentGpsKm  = [...agg.values()].reduce((s, v) => s + v.km, 0);

            if (currentGpsKm > 0 && Math.abs(expectedGpsKm - currentGpsKm) > 0.01) {
                const ratio = expectedGpsKm / currentGpsKm;
                for (const [k, v] of agg) {
                    if (v.km > 0) {
                        v.km = Math.round(v.km * ratio * 100) / 100;
                    }
                }

                // Fix rounding diff on the day with the most km
                const newTotal = [...agg.values()].reduce((s, v) => s + v.km, 0);
                const diff = Math.round((expectedGpsKm - newTotal) * 100) / 100;
                if (Math.abs(diff) > 0) {
                    let maxDay = null;
                    let maxKm  = -Infinity;
                    for (const [k, v] of agg) {
                        if (v.km > 0 && v.km > maxKm) { maxKm = v.km; maxDay = k; }
                    }
                    if (maxDay !== null) {
                        agg.get(maxDay).km = Math.round((agg.get(maxDay).km + diff) * 100) / 100;
                    }
                }
            }
        }

        // Parse refuel dates
        const year  = dateFrom ? parseInt(dateFrom.slice(0, 4), 10) : new Date().getFullYear();
        const month = dateFrom ? parseInt(dateFrom.slice(5, 7), 10) : new Date().getMonth() + 1;
        const refuelSet = new Set();

        for (const part of refuelRaw.split(/[,;\s]+/)) {
            const p = part.trim();
            if (!p) continue;
            const m = /^(\d{1,2})[.\-\/](\d{1,2})(?:[.\-\/](\d{2,4}))?/.exec(p);
            if (m) {
                try {
                    const d  = parseInt(m[1], 10);
                    const mo = parseInt(m[2], 10);
                    let   y  = m[3] ? parseInt(m[3], 10) : year;
                    if (y < 100) y += 2000;
                    const padM = String(mo).padStart(2, '0');
                    const padD = String(d).padStart(2, '0');
                    refuelSet.add(`${y}-${padM}-${padD}`);
                } catch (e) {
                    // ignore invalid dates
                }
            }
        }

        const xlsBuf = await generateExcel(
            plate, carModel, dateFrom, dateTo,
            driver, odometer, refuelSet, agg,
            tripPurpose
        );

        const baseName = path.basename(filename, path.extname(filename));
        const outName  = `${baseName}.xlsx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
        res.send(xlsBuf);

    } catch (err) {
        console.error(err);
        res.status(500).json({ error: String(err.message || err), detail: err.stack || '' });
    }
});

// ── Start ─────────────────────────────────────────────────────────────────────

const PORT = parseInt(process.env.PORT, 10) || 8080;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n\x1b[32m✔ Server running!\x1b[0m`);
    console.log(`\x1b[36m➜ Local:\x1b[0m    http://localhost:${PORT}`);
    console.log(`\x1b[36m➜ Network:\x1b[0m  http://0.0.0.0:${PORT} (Docker)\n`);
});
