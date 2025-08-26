/**
 * =========================================================
 * RETENTION APP — server side (Google Apps Script) - FIXED
 * =========================================================
 */

const ID_HEADER = 'id';
const TAGS_HEADER = 'tags';
const QTAG_HEADER = 'qtag';

/* -------------------------- MENU -------------------------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Retention')
    .addItem('Open app', 'openRetentionApp')
    .addSeparator()
    .addItem('Install warmup triggers', 'installWarmTriggers')
    .addItem('Warm up now', 'warmAll')
    .addItem('Clear cache', 'clearAllCaches')
    .addItem('Debug data', 'debugRetentionData')
    .addToUi();
}

function openRetentionApp() {
  const html = HtmlService.createTemplateFromFile('App')
    .evaluate()
    .setWidth(1200)
    .setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'Retention Rate');
}

/* ------------------------- VIP INFO ------------------------ */
function getVipInfo() {
  const sh = SpreadsheetApp.getActive().getSheetByName('VIP INFO');
  const res = { vip: [], slvip: [] };
  if (!sh) return res;
  
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][0] || '').trim();
    if (!name) continue;
    
    res.vip.push({
      name,
      tags: tokenizeTags_(values[i][1])
    });
    res.slvip.push({
      name,
      tags: tokenizeTags_(values[i][2])
    });
  }
  return res;
}

/* -------------------------- CACHE -------------------------- */
function cacheFile_() {
  const ss = SpreadsheetApp.getActive();
  const fname = 'RetentionCache_' + ss.getId() + '.json';
  const file = DriveApp.getFileById(ss.getId());
  const folder = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
  
  const it = folder.getFilesByName(fname);
  if (it.hasNext()) return it.next();
  
  return folder.createFile(fname, '{}', MimeType.PLAIN_TEXT);
}

function readCache_() {
  try {
    const content = cacheFile_().getBlob().getDataAsString('utf-8') || '{}';
    return JSON.parse(content);
  } catch (e) {
    console.error('Cache read error:', e);
    return {};
  }
}

function writeCache_(obj) {
  try {
    cacheFile_().setContent(JSON.stringify(obj));
    return obj;
  } catch (e) {
    console.error('Cache write error:', e);
    return obj;
  }
}

function clearCache() {
  try {
    writeCache_({});
    SpreadsheetApp.getUi().alert('Cache cleared successfully!');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error clearing cache: ' + e.message);
  }
}
function clearPropertiesCache() {
  try {
    const props = PropertiesService.getScriptProperties();
    const keys = props.getKeys();
    keys.forEach(key => {
      if (key.includes('_cache') || key.includes('sheets_')) {
        props.deleteProperty(key);
      }
    });
    console.log('Properties cache cleared');
  } catch (e) {
    console.error('Error clearing properties cache:', e);
  }
}

function clearAllCaches() {
  clearCache();
  clearPropertiesCache();
  SpreadsheetApp.getUi().alert('All caches cleared successfully!');
}

/* ----------------------- DEBUGGING ------------------------- */
function debugRetentionData() {
  const ss = SpreadsheetApp.getActive();
  const months = MONTHS_();
  const sheets = ss.getSheets();
  
  console.log('=== DEBUGGING RETENTION DATA ===');
  
  let ftd = [], active = [];
  
  for (const sh of sheets) {
    const name = sh.getName();
    const low = name.toLowerCase();
    const idx = months.findIndex(m => low.includes(m.toLowerCase()));
    
    if (idx === -1) continue;
    
    if (low.includes('ftd')) {
      const rowCount = sh.getLastRow();
      ftd.push({ idx, label: months[idx], sheet: sh, name, rowCount });
      console.log(`FTD: ${name} (${months[idx]}) - idx: ${idx}, rows: ${rowCount}`);
    }
    if (low.includes('active')) {
      const rowCount = sh.getLastRow();
      active.push({ idx, label: months[idx], sheet: sh, name, rowCount });
      console.log(`ACTIVE: ${name} (${months[idx]}) - idx: ${idx}, rows: ${rowCount}`);
    }
  }
  
  ftd.sort((a, b) => a.idx - b.idx);
  active.sort((a, b) => a.idx - b.idx);
  
  console.log('=== FTD SHEETS ===');
  ftd.forEach(f => console.log(`${f.label} (idx: ${f.idx}) - ${f.rowCount} rows`));
  
  console.log('=== ACTIVE SHEETS ===');
  active.forEach(a => console.log(`${a.label} (idx: ${a.idx}) - ${a.rowCount} rows`));
  
  const augustActive = active.find(a => a.label === 'August');
  if (augustActive) {
    console.log('=== AUGUST ACTIVE DETAILS ===');
    const sheet = augustActive.sheet;
    const values = sheet.getDataRange().getValues();
    console.log(`Total rows: ${values.length}`);
    console.log(`Headers: ${values[0]}`);
    
    let validIds = 0;
    for (let i = 1; i < values.length; i++) {
      const id = String(values[i][0] || '').trim().toLowerCase();
      if (id.includes('bets:') || id.includes('betsio:')) {
        validIds++;
      }
    }
    console.log(`Valid brand IDs in August: ${validIds}`);
  } else {
    console.log('August ACTIVE sheet not found!');
  }
  
  SpreadsheetApp.getUi().alert('Debug completed. Check logs in Apps Script editor.');
}

/* --------------------- WARMUP TRIGGERS --------------------- */
function installWarmTriggers() {
  const ssId = SpreadsheetApp.getActive().getId();
  
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === 'warmAfterChange' || f === 'timeWarmHourly')
      ScriptApp.deleteTrigger(t);
  });
  
  ScriptApp.newTrigger('warmAfterChange').forSpreadsheet(ssId).onChange().create();
  ScriptApp.newTrigger('timeWarmHourly').timeBased().everyHours(1).create();
  
  SpreadsheetApp.getUi().alert('Warmup triggers installed.');
}

function warmAfterChange() {
  Utilities.sleep(2000);
  warmAll();
}

function timeWarmHourly() {
  warmAll();
}

/* ------------------------ PRE-WARM ALL --------------------- */
function warmAll() {
  try {
    console.log('Starting warmAll...');
    
    const vipInfo = getVipInfo();
    const qmapDump = dumpQtagMap_();
    const projects = ['bets', 'betsio'];
    const segments = ['general', 'organic', 'partners', 'vip', 'slvip'];
    
    // Збираємо всі унікальні qtag номери для partners
    const uniqueQtags = new Set();
    const qmap = loadQtagMap_(SpreadsheetApp.getActive());
    qmap.forEach((value, key) => {
      const parts = key.split('|');
      if (parts[1]) uniqueQtags.add(parts[1]);
    });
    const qtagNumbers = Array.from(uniqueQtags).slice(0, 10); // Обмежуємо 10 найпопулярнішими
    
    // Збираємо всі унікальні qtag з General Info
const allQtags = [];
const giSheet = SpreadsheetApp.getActive().getSheetByName('General Info');
if (giSheet) {
  const giValues = giSheet.getDataRange().getValues();
  for (let i = 1; i < giValues.length; i++) {
    const num = String(giValues[i][1] || '').trim();
    if (num && !allQtags.includes(num)) {
      allQtags.push(num);
    }
  }
}

const subs = {
  general: ['all'],
  organic: ['all'],
  partners: ['all', 'partners', 'streamers', 'cross-sell'],
  vip: ['all', ...vipInfo.vip.map(v => v.name.toLowerCase())],
  slvip: ['all', ...vipInfo.slvip.map(v => v.name.toLowerCase())]
};
    
    const data = {};
    let processedCount = 0;
    const totalCount = projects.length * segments.reduce((sum, s) => sum + subs[s].length, 0);
    
    for (const p of projects) {
  data[p] = {};
  for (const s of segments) {
    data[p][s] = {};
    for (const sub of subs[s]) {
      try {
        console.log(`Processing: ${p} / ${s} / ${sub}`);
        
        // Для partners кешуємо також з qtag
        if (s === 'partners' && allQtags.length > 0) {
          // Кешуємо без qtag
          data[p][s][sub] = data[p][s][sub] || {};
          data[p][s][sub][''] = getMatrixData({
            project: p,
            segment: s,
            sub
          });
          
          // Кешуємо топ-5 qtag
          const topQtags = allQtags.slice(0, 5);
          for (const qtag of topQtags) {
            console.log(`Processing with qtag: ${qtag}`);
            data[p][s][sub][qtag] = getMatrixData({
              project: p,
              segment: s,
              sub,
              qtagNumber: qtag
            });
          }
        } else {
          data[p][s][sub] = getMatrixData({
            project: p,
            segment: s,
            sub
          });
        }
            processedCount++;
            
            if (processedCount % 5 === 0) {
              console.log(`Progress: ${processedCount}/${totalCount}`);
            }
          } catch (e) {
            console.error(`Error processing ${p}/${s}/${sub}:`, e);
            data[p][s][sub] = { error: e.message };
          }
        }
      }
    }
    
    const payload = {
      generatedAt: new Date().toISOString(),
      fileUpdatedAt: DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getLastUpdated().getTime(),
      vipInfo,
      qmapDump,
      data
    };
    
    writeCache_(payload);
    console.log('warmAll completed successfully');
    return { ok: true, processed: processedCount };
    
  } catch (e) {
    console.error('warmAll error:', e);
    return { error: e.message };
  }
}

function getPrewarmed() {
  // Не блокуємо UI: повертаємо те, що є, а прогрів запускаємо у фоні
  const cached = readCache_();
  const now = Date.now();
  const maxAge = 2 * 60 * 60 * 1000; // 2 години

  // Якщо кеш свіжий — просто віддаємо
  if (cached && cached.generatedAt) {
    const age = now - new Date(cached.generatedAt).getTime();
    if (age <= maxAge) return cached;
  }

  // Якщо кеш застарілий або відсутній — запускаємо фон
  try {
    // спроба поставити разовий тригер на найближчі кілька секунд
    ScriptApp.newTrigger('warmAll').timeBased().after(3 * 1000).create();
  } catch (e) {
    // якщо не вдалося — просто запускаємо без очікування (ми все одно нічого не чекаємо)
    try { warmAll(); } catch (_) {}
  }

  // Повертаємо те, що є (може бути пусто) — фронт сам підтягне живі дані
  return cached || {};
}

function quickPreload() {
  try {
    const ss = SpreadsheetApp.getActive();
    const months = MONTHS_();
    const sheets = ss.getSheets();
    
    // Швидко збираємо базову інформацію
    const quickData = {
      ftdMonths: [],
      activeMonths: [],
      totalRows: 0
    };
    
    for (const sh of sheets) {
      const name = sh.getName();
      const low = name.toLowerCase();
      const idx = months.findIndex(m => low.includes(m.toLowerCase()));
      
      if (idx !== -1) {
        if (low.includes('ftd')) {
          quickData.ftdMonths.push(months[idx]);
          quickData.totalRows += sh.getLastRow();
        }
        if (low.includes('active')) {
          quickData.activeMonths.push(months[idx]);
          quickData.totalRows += sh.getLastRow();
        }
      }
    }
    
    return quickData;
  } catch (e) {
    console.error('quickPreload error:', e);
    return null;
  }
}
/* ----------------------- ONE MATRIX ------------------------ */
function getMatrixData(options) {
  try {
    options = options || {};
    const project = normProjectKey_(options.project || 'bets');
    const segment = String(options.segment || 'general').toLowerCase();
    const sub = String(options.sub || 'all').toLowerCase();
    const qtagNumber = options.qtagNumber ? String(options.qtagNumber) : null;
    
    console.log(`getMatrixData: ${project}/${segment}/${sub}${qtagNumber ? '/' + qtagNumber : ''}`);
    
    const ss = SpreadsheetApp.getActive();
    const months = MONTHS_();
    const sheets = ss.getSheets();
    
    let ftd = [], active = [];
    
    // Пошук аркушів
    // Кешуємо аркуші в Properties для швидкого доступу
const cacheKey = 'sheets_cache';
let sheetsCache;
try {
  const cached = PropertiesService.getScriptProperties().getProperty(cacheKey);
  sheetsCache = cached ? JSON.parse(cached) : null;
} catch(e) {}

if (!sheetsCache) {
  sheetsCache = { ftd: [], active: [] };
  for (const sh of sheets) {
    const name = sh.getName();
    const low = name.toLowerCase();
    const idx = months.findIndex(m => low.includes(m.toLowerCase()));
    
    if (idx === -1) continue;
    
    if (low.includes('ftd')) {
      sheetsCache.ftd.push({ idx, label: months[idx], name });
    }
    if (low.includes('active')) {
      sheetsCache.active.push({ idx, label: months[idx], name });
    }
  }
  
  try {
    PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(sheetsCache));
  } catch(e) {}
}

// Використовуємо кешовані дані
for (const f of sheetsCache.ftd) {
  ftd.push({
    ...f,
    sheet: ss.getSheetByName(f.name)
  });
}
for (const a of sheetsCache.active) {
  active.push({
    ...a,
    sheet: ss.getSheetByName(a.name)
  });
}
    
    console.log(`Found ${ftd.length} FTD sheets, ${active.length} ACTIVE sheets`);
    
    if (!ftd.length || !active.length) {
      return { error: 'No FTD/ACTIVE sheets found.' };
    }
    
    ftd.sort((a, b) => a.idx - b.idx);
    active.sort((a, b) => a.idx - b.idx);
    
    if (Array.isArray(options.includeFTD) && options.includeFTD.length) {
      const set = new Set(options.includeFTD);
      ftd = ftd.filter(x => set.has(x.label));
    }
    if (Array.isArray(options.includeACTIVE) && options.includeACTIVE.length) {
      const set = new Set(options.includeACTIVE);
      active = active.filter(x => set.has(x.label));
    }
    
    const qmap = loadQtagMap_(ss);
    const vipInfo = getVipInfo();
    const allSets = buildVipSets_(vipInfo);
    
    const actSheetsDesc = collectActiveSheets_(ss).sort((a, b) => b.idx - a.idx);
    const latestActiveMap = aggregateActiveMap_(actSheetsDesc);
    
    const ftdSets = new Map();
    for (const f of ftd) {
      try {
        const ids = collectIds_(
          f.sheet, project, segment, sub, qmap,
          allSets.vipNameToSet, allSets.slvipNameToSet,
          allSets.vipTagSet, allSets.slvipTagSet,
          latestActiveMap, qtagNumber
        );
        ftdSets.set(f.label, ids);
        console.log(`FTD ${f.label}: ${ids.size} IDs`);
      } catch (e) {
        console.error(`Error processing FTD sheet ${f.name}:`, e);
        ftdSets.set(f.label, new Set());
      }
    }
    
    const actSets = new Map();
    for (const a of active) {
      try {
        const ids = collectIds_(
          a.sheet, project, segment, sub, qmap,
          allSets.vipNameToSet, allSets.slvipNameToSet,
          allSets.vipTagSet, allSets.slvipTagSet,
          latestActiveMap, qtagNumber
        );
        actSets.set(a.label, ids);
        console.log(`ACTIVE ${a.label}: ${ids.size} IDs`);
      } catch (e) {
        console.error(`Error processing ACTIVE sheet ${a.name}:`, e);
        actSets.set(a.label, new Set());
      }
    }
    
    const totals = new Map();
    for (const f of ftd) {
      totals.set(f.label, ftdSets.get(f.label).size);
    }
    
    let firstNonZero = ftd.findIndex(f => (totals.get(f.label) || 0) > 0);
    if (firstNonZero === -1) firstNonZero = 0;
    ftd = ftd.slice(firstNonZero);
    
    let rows = [];
    let maxPct = 0;
    
    // ВИПРАВЛЕНА логіка обчислення
    for (const f of ftd) {
      const total = totals.get(f.label) || 0;
      const cells = [];
      
      for (const a of active) {
        // ВИПРАВЛЕННЯ: тепер правильно порівнюємо індекси
        // ACTIVE має бути ПІСЛЯ FTD для retention аналізу
        if (a.idx <= f.idx) {
          cells.push({
            alabel: a.label,
            count: null,
            pct: null
          });
          continue;
        }
        
        const ftdSet = ftdSets.get(f.label);
        const actSet = actSets.get(a.label);
        
        if (!ftdSet || !actSet) {
          console.log(`Warning: Missing sets for ${f.label} -> ${a.label}`);
          cells.push({
            alabel: a.label,
            count: 0,
            pct: 0
          });
          continue;
        }
        
        const cnt = intersectCount_(ftdSet, actSet);
        const pct = total ? (cnt / total) : 0;
        
        console.log(`${f.label}(${total}) -> ${a.label}: ${cnt} (${(pct*100).toFixed(2)}%)`);
        
        if (pct > maxPct) maxPct = pct;
        
        cells.push({
          alabel: a.label,
          count: cnt,
          pct
        });
      }
      
      rows.push({
        ftdLabel: f.label,
        total,
        cells
      });
    }
    
    const keepCol = active.map((_, j) => rows.some(r => r.cells[j] && r.cells[j].count !== null));
    const active2 = active.filter((_, j) => keepCol[j]);
    rows = rows.map(r => ({
      ...r,
      cells: r.cells.filter((_, j) => keepCol[j])
    }));
    
    let kpi = null;
    if (rows.length && active2.length) {
      const lastActive = active2[active2.length - 1];
      const prevFtd = ftd.slice().reverse().find(f => f.idx < lastActive.idx);
      
      if (prevFtd) {
        const row = rows.find(r => r.ftdLabel === prevFtd.label);
        const c = row ? row.cells.find(x => x.alabel === lastActive.label) : null;
        const overlap = c ? (c.count || 0) : 0;
        const total = row ? (row.total || 0) : 0;
        const pct = total ? (overlap / total) : 0;
        
        kpi = {
          from: prevFtd.label,
          to: lastActive.label,
          total,
          overlap,
          pct
        };
      }
    }
    
    return {
      meta: {
        project,
        segment,
        sub,
        ftdHeader: ftd.map(x => x.label),
        activeHeader: active2.map(x => x.label),
        maxPct,
        kpi,
        vipSubtabs: vipInfo.vip.map(v => v.name),
        slvipSubtabs: vipInfo.slvip.map(v => v.name)
      },
      rows
    };
    
  } catch (e) {
    console.error('getMatrixData error:', e);
    return { error: e.message };
  }
}

/* --------------- MODALS: PLAYERS LISTS -------------------- */
function getCellPlayers(opts) {
 try {
   const project = normProjectKey_(opts && opts.project || 'bets');
   const segment = String(opts && opts.segment || 'general').toLowerCase();
   const sub = String(opts && opts.sub || 'all').toLowerCase();
   const ftdLabel = String(opts && opts.ftdLabel || '');
   const activeLabel = String(opts && opts.activeLabel || '');
   const listType = String(opts && opts.listType || 'retention').toLowerCase();
   const qtagNumber = opts && opts.qtagNumber ? String(opts.qtagNumber) : null;
   
   const ss = SpreadsheetApp.getActive();
   const shFTD = findSheet_(ss, ftdLabel, 'ftd');
   const shACT = findSheet_(ss, activeLabel, 'active');
   
   if (!shFTD || !shACT) {
     return { error: 'Sheets not found.' };
   }
   
   const qmap = loadQtagMap_(ss);
   const vipInf = getVipInfo();
   const allSets = buildVipSets_(vipInf);
   const latestActiveMap = aggregateActiveMap_(collectActiveSheets_(ss).sort((a, b) => b.idx - a.idx));
   
   const setF = collectIds_(shFTD, project, segment, sub, qmap, allSets.vipNameToSet, allSets.slvipNameToSet, allSets.vipTagSet, allSets.slvipTagSet, latestActiveMap, qtagNumber);
   const setA = collectIds_(shACT, project, segment, sub, qmap, allSets.vipNameToSet, allSets.slvipNameToSet, allSets.vipTagSet, allSets.slvipTagSet, latestActiveMap, qtagNumber);
   
   let kept = new Set();
   if (listType === 'retention') {
     for (const id of setF) {
       if (setA.has(id)) kept.add(id);
     }
   } else {
     for (const id of setF) {
       if (!setA.has(id)) kept.add(id);
     }
   }
   
   const mapA = mapRowsById_(shACT);
   const mapF = mapRowsById_(shFTD);
   const players = [];
   
   for (const id of kept) {
     const prefer = mapA.get(id);
     const alt1 = latestActiveMap.get(id);
     const alt2 = mapF.get(id);
     const best = _buildBestRow_(prefer, [alt1, alt2]);
     
     if (!best || !best.id) continue;
     
     const info = extractRecordFlexible_(best, project);
     let manager = '';
     
     if ((segment === 'vip' || segment === 'slvip') && sub === 'all') {
       const tokens = tokenizeTags_(best.tags);
       manager = (segment === 'vip') ? detectManager_(tokens, vipInf.vip) : detectManager_(tokens, vipInf.slvip);
     }
     
     info.manager = manager || '';
     players.push(info);
   }
   
   let managers = [];
   if ((segment === 'vip' || segment === 'slvip') && sub === 'all') {
     managers = (segment === 'vip' ? vipInf.vip : vipInf.slvip).map(v => v.name);
   }
   
   players.sort((a, b) => {
     const da = a.created_ts || 0, db = b.created_ts || 0;
     if (db !== da) return db - da;
     return String(a.email || '').localeCompare(String(b.email || ''));
   });
   
   return {
     mode: 'cell',
     ftd: ftdLabel,
     active: activeLabel,
     listType,
     count: players.length,
     players,
     managers
   };
   
 } catch (e) {
   console.error('getCellPlayers error:', e);
   return { error: e.message };
 }
}

function getFtdPlayers(opts) {
 try {
   const project = normProjectKey_(opts && opts.project || 'bets');
   const segment = String(opts && opts.segment || 'general').toLowerCase();
   const sub = String(opts && opts.sub || 'all').toLowerCase();
   const ftdLabel = String(opts && opts.ftdLabel || '');
   const qtagNumber = opts && opts.qtagNumber ? String(opts.qtagNumber) : null;
   
   const ss = SpreadsheetApp.getActive();
   const shFTD = findSheet_(ss, ftdLabel, 'ftd');
   
   if (!shFTD) {
     return { error: 'FTD sheet not found.' };
   }
   
   const qmap = loadQtagMap_(ss);
   const vipInf = getVipInfo();
   const allSets = buildVipSets_(vipInf);
   const latestActiveMap = aggregateActiveMap_(collectActiveSheets_(ss).sort((a, b) => b.idx - a.idx));
   
   const setF = collectIds_(shFTD, project, segment, sub, qmap, allSets.vipNameToSet, allSets.slvipNameToSet, allSets.vipTagSet, allSets.slvipTagSet, latestActiveMap, qtagNumber);
   
   const players = [];
   const mapF = mapRowsById_(shFTD);
   
   for (const id of setF) {
     const prefer = latestActiveMap.get(id);
     const alt1 = mapF.get(id);
     const best = _buildBestRow_(prefer, [alt1]);
     
     if (!best || !best.id) continue;
     
     const info = extractRecordFlexible_(best, project);
     let manager = '';
     
     if ((segment === 'vip' || segment === 'slvip') && sub === 'all') {
       const tokens = tokenizeTags_(best.tags);
       manager = (segment === 'vip') ? detectManager_(tokens, vipInf.vip) : detectManager_(tokens, vipInf.slvip);
     }
     
     info.manager = manager || '';
     players.push(info);
   }
   
   let managers = [];
   if ((segment === 'vip' || segment === 'slvip') && sub === 'all') {
     managers = (segment === 'vip' ? vipInf.vip : vipInf.slvip).map(v => v.name);
   }
   
   players.sort((a, b) => {
     const da = a.created_ts || 0, db = b.created_ts || 0;
     if (db !== da) return db - da;
     return String(a.email || '').localeCompare(String(b.email || ''));
   });
   
   return {
     mode: 'ftd',
     ftd: ftdLabel,
     count: players.length,
     players,
     managers
   };
   
 } catch (e) {
   console.error('getFtdPlayers error:', e);
   return { error: e.message };
 }
}

function getColumnPlayers(opts) {
 try {
   const project = normProjectKey_(opts && opts.project || 'bets');
   const segment = String(opts && opts.segment || 'general').toLowerCase();
   const sub = String(opts && opts.sub || 'all').toLowerCase();
   const activeLabel = String(opts && opts.activeLabel || '');
   const kind = String(opts && opts.kind || 'retention').toLowerCase();
   const qtagNumber = opts && opts.qtagNumber ? String(opts.qtagNumber) : null;
   
   const ss = SpreadsheetApp.getActive();
   const shACT = findSheet_(ss, activeLabel, 'active');
   
   if (!shACT) {
     return { error: 'ACTIVE sheet not found.' };
   }
   
   const months = MONTHS_();
   const sheets = ss.getSheets();
   
   const ftdSheets = sheets
     .filter(sh => {
       const low = sh.getName().toLowerCase();
       return low.includes('ftd') && months.some(m => low.includes(m.toLowerCase()));
     })
     .map(sh => {
       const low = sh.getName().toLowerCase();
       const idx = months.findIndex(m => low.includes(m.toLowerCase()));
       return { idx, label: months[idx], sheet: sh };
     })
     .sort((a, b) => a.idx - b.idx);
   
   const qmap = loadQtagMap_(ss);
   const vipInf = getVipInfo();
   const allSets = buildVipSets_(vipInf);
   
   const allActiveAsc = collectActiveSheets_(ss).sort((a, b) => a.idx - b.idx);
   const allActiveDesc = allActiveAsc.slice().sort((a, b) => b.idx - a.idx);
   const latestActiveMap = aggregateActiveMap_(allActiveDesc);
   
   const target = allActiveAsc.find(x => x.label === activeLabel);
   if (!target) {
     return { error: 'Target ACTIVE month not found.' };
   }
   
   if (kind === 'churn') {
     const activeMaps = allActiveAsc.map(a => ({
       idx: a.idx,
       label: a.label,
       set: collectIds_(a.sheet, project, segment, sub, qmap, allSets.vipNameToSet, allSets.slvipNameToSet, allSets.vipTagSet, allSets.slvipTagSet, latestActiveMap, qtagNumber),
       rowMap: mapRowsById_(a.sheet)
     }));
     
     const targetSet = activeMaps.find(a => a.label === activeLabel).set;
     const byActive = {};
     const order = [];
     
     function push(lbl, info) {
       if ((segment === 'vip' || segment === 'slvip')) {
         if (!byActive[lbl]) {
           byActive[lbl] = {};
           order.push(lbl);
         }
         const m = info.manager || '-';
         (byActive[lbl][m] = byActive[lbl][m] || []).push(info);
       } else {
         if (!byActive[lbl]) {
           byActive[lbl] = [];
           order.push(lbl);
         }
         byActive[lbl].push(info);
       }
     }
     
     const ftdRowMaps = {};
     ftdSheets.forEach(f => ftdRowMaps[f.label] = mapRowsById_(f.sheet));
     
     for (const f of ftdSheets) {
       if (f.idx >= target.idx) continue;
       
       const setF = collectIds_(f.sheet, project, segment, sub, qmap, allSets.vipNameToSet, allSets.slvipNameToSet, allSets.vipTagSet, allSets.slvipTagSet, latestActiveMap, qtagNumber);
       
       for (const id of setF) {
         if (targetSet.has(id)) continue;
         
         let lastLabel = null, lastRow = null;
         for (let i = activeMaps.length - 1; i >= 0; i--) {
           const a = activeMaps[i];
           if (a.idx <= f.idx) continue;
           if (a.idx >= target.idx) continue;
           if (a.set.has(id)) {
             lastLabel = a.label;
             lastRow = a.rowMap.get(id) || latestActiveMap.get(id);
             break;
           }
         }
         
         if (!lastLabel) {
           lastLabel = f.label;
           lastRow = ftdRowMaps[f.label].get(id) || latestActiveMap.get(id);
         }
         
         if (!lastRow) continue;
         
         const best = _buildBestRow_(lastRow, [ftdRowMaps[f.label].get(id), latestActiveMap.get(id)]);
         const info = extractRecordFlexible_(best, project);
         
         if ((segment === 'vip' || segment === 'slvip')) {
           const tokens = tokenizeTags_(best.tags);
           info.manager = (segment === 'vip' ? detectManager_(tokens, vipInf.vip) : detectManager_(tokens, vipInf.slvip)) || '';
         }
         
         push(`ACTIVE IN ${lastLabel}`, info);
       }
     }
     
     const monthOrder = allActiveAsc.map(x => `ACTIVE IN ${x.label}`);
     const finalOrder = [];
     const seen = new Set();
     
     monthOrder.forEach(lbl => {
       if (byActive[lbl] && !seen.has(lbl)) {
         finalOrder.push(lbl);
         seen.add(lbl);
       }
     });
     
     return {
       mode: 'column-churn',
       byActive,
       order: finalOrder
     };
   }
   
   const setA = collectIds_(shACT, project, segment, sub, qmap, allSets.vipNameToSet, allSets.slvipNameToSet, allSets.vipTagSet, allSets.slvipTagSet, latestActiveMap, qtagNumber);
   const mapA = mapRowsById_(shACT);
   
   const monthsOrder = [];
   const byMonth = {};
   
   for (const f of ftdSheets) {
     if (f.idx >= target.idx) continue;
     
     const setF = collectIds_(f.sheet, project, segment, sub, qmap, allSets.vipNameToSet, allSets.slvipNameToSet, allSets.vipTagSet, allSets.slvipTagSet, latestActiveMap, qtagNumber);
     
     let any = false;
     for (const id of setF) {
       if (!setA.has(id)) continue;
       
       const prefer = mapA.get(id);
       const alt1 = latestActiveMap.get(id);
       const best = _buildBestRow_(prefer, [alt1, null]);
       
       if (!best || !best.id) continue;
       
       const info = extractRecordFlexible_(best, project);
       let manager = '';
       
       if ((segment === 'vip' || segment === 'slvip') && sub === 'all') {
         const tokens = tokenizeTags_(best.tags);
         manager = (segment === 'vip') ? detectManager_(tokens, vipInf.vip) : detectManager_(tokens, vipInf.slvip);
       }
       
       info.manager = manager || '';
       
       if (!byMonth[f.label]) {
         byMonth[f.label] = {};
         monthsOrder.push(f.label);
       }
       
       const key = ((segment === 'vip' || segment === 'slvip') && sub === 'all') ? (manager || '-') : 'Players';
       (byMonth[f.label][key] = byMonth[f.label][key] || []).push(info);
       any = true;
     }
     
     if (!any) continue;
   }
   
   return {
     mode: 'column',
     months: monthsOrder,
     byMonth
   };
   
 } catch (e) {
   console.error('getColumnPlayers error:', e);
   return { error: e.message };
 }
}

/* -------------------------- EXPORT ------------------------- */
function exportPlayersNewSpreadsheet(payload) {
 try {
   const title = payload && payload.title || ('Players_' + new Date().toISOString().slice(0, 10));
   const headers = payload && payload.headers || defaultHeaders_();
   
   const ss = SpreadsheetApp.create(title);
   const sh = ss.getActiveSheet();
   
   if (payload && payload.structure) {
     writeStructuredTableBatched_(sh, headers, payload.structure);
     SpreadsheetApp.flush(); // Форсуємо збереження
     applyGreenTableTheme_(sh, headers.length, false);
   } else {
     writeTableBatched_(sh, headers, (payload && payload.players) || []);
     SpreadsheetApp.flush(); // Форсуємо збереження
     applyGreenTableTheme_(sh, headers.length, true);
   }
   
   return {
     url: ss.getUrl(),
     id: ss.getId(),
     sheetName: sh.getName()
   };
 } catch (e) {
   return { error: String(e && e.message || e) };
 }
}

function exportPlayersExistingSpreadsheet(payload) {
 try {
   const spreadsheetId = payload && payload.spreadsheetId;
   if (!spreadsheetId) throw new Error('Missing spreadsheetId');
   
   const sheetName = payload && payload.sheetName || 'Players';
   const headers = payload && payload.headers || defaultHeaders_();
   
   const ss = SpreadsheetApp.openById(spreadsheetId);
   let sh = ss.getSheetByName(sheetName);
   if (!sh) sh = ss.insertSheet(sheetName);
   
   sh.clear();
   
   if (payload && payload.structure) {
     writeStructuredTable_(sh, headers, payload.structure);
     applyGreenTableTheme_(sh, headers.length, false);
   } else {
     writeTable_(sh, headers, (payload && payload.players) || []);
     applyGreenTableTheme_(sh, headers.length, true);
   }
   
   return {
     url: ss.getUrl(),
     id: ss.getId(),
     sheetName: sh.getName()
   };
 } catch (e) {
   return { error: String(e && e.message || e) };
 }
}

function writeTable_(sh, headers, players) {
 sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
 
 if (players.length) {
   const rows = players.map(p => toRow_(p, headers));
   sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
 }
 
 sh.autoResizeColumns(1, headers.length);
}

function writeTableBatched_(sh, headers, players) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  
  if (players.length) {
    const BATCH_SIZE = 500; // Записуємо по 500 рядків
    const rows = players.map(p => toRow_(p, headers));
    
    for (let i = 0; i < rows.length; i += BATCH_SIZE) {
      const batch = rows.slice(i, Math.min(i + BATCH_SIZE, rows.length));
      sh.getRange(i + 2, 1, batch.length, headers.length).setValues(batch);
      
      // Даємо час на обробку кожної пачки
      if (i + BATCH_SIZE < rows.length) {
        Utilities.sleep(100);
      }
    }
  }
  
  sh.autoResizeColumns(1, headers.length);
}
function writeStructuredTable_(sh, headers, structure) {
 let r = 1;
 sh.getRange(r, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
 r++;

 function writeStructuredTableBatched_(sh, headers, structure) {
  const BATCH_SIZE = 500;
  let allRows = [];
  let rowFormats = []; // Для збереження інформації про форматування
  
  // Збираємо всі рядки в один масив
  allRows.push(headers);
  rowFormats.push({type: 'header'});
  
  structure.forEach(mon => {
    allRows.push([mon.title]);
    rowFormats.push({type: 'section', colspan: headers.length});
    
    (mon.groups || []).forEach(g => {
      allRows.push(['  ' + g.title]);
      rowFormats.push({type: 'subsection', colspan: headers.length});
      
      (g.players || []).forEach(p => {
        allRows.push(toRow_(p, headers));
        rowFormats.push({type: 'data'});
      });
    });
  });
  
  // Записуємо пакетами
  let currentRow = 1;
  for (let i = 0; i < allRows.length; i += BATCH_SIZE) {
    const batch = allRows.slice(i, Math.min(i + BATCH_SIZE, allRows.length));
    const formats = rowFormats.slice(i, Math.min(i + BATCH_SIZE, allRows.length));
    
    // Записуємо дані
    batch.forEach((row, idx) => {
      const format = formats[idx];
      if (format.type === 'header') {
        sh.getRange(currentRow, 1, 1, headers.length).setValues([row]).setFontWeight('bold');
      } else if (format.type === 'section' || format.type === 'subsection') {
        const range = sh.getRange(currentRow, 1, 1, headers.length);
        range.mergeAcross().setValue(row[0]);
        
        if (format.type === 'section') {
          range.setFontWeight('bold').setBackground('#f1f5f9')
            .setBorder(true, true, false, true, false, false, '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID);
        } else {
          range.setFontWeight('bold').setBackground('#eef2ff')
            .setBorder(false, true, false, true, false, false, '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID);
        }
      } else {
        sh.getRange(currentRow, 1, 1, headers.length).setValues([row]);
      }
      currentRow++;
    });
    
    if (i + BATCH_SIZE < allRows.length) {
      Utilities.sleep(100);
      SpreadsheetApp.flush();
    }
  }
  
  sh.autoResizeColumns(1, headers.length);
}
 
 structure.forEach(mon => {
   sh.getRange(r, 1, 1, headers.length).mergeAcross().setValue(mon.title)
     .setFontWeight('bold').setBackground('#f1f5f9')
     .setBorder(true, true, false, true, false, false, '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID);
   r++;
   
   (mon.groups || []).forEach(g => {
     sh.getRange(r, 1, 1, headers.length).mergeAcross().setValue('  ' + g.title)
       .setFontWeight('bold').setBackground('#eef2ff')
       .setBorder(false, true, false, true, false, false, '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID);
     r++;
     
     const rows = (g.players || []).map(p => toRow_(p, headers));
     if (rows.length) {
       sh.getRange(r, 1, rows.length, headers.length).setValues(rows);
       r += rows.length;
     }
   });
 });
 
 sh.autoResizeColumns(1, headers.length);
}

function applyGreenTableTheme_(sh, cols, zebra) {
 const last = sh.getLastRow() || 1;
 sh.getRange(1, 1, 1, cols).setBackground('#16a34a').setFontColor('#ffffff').setFontWeight('bold');
 
 const body = sh.getRange(2, 1, Math.max(0, last - 1), cols);
 body.setBorder(true, true, true, true, true, true, '#d1fae5', SpreadsheetApp.BorderStyle.SOLID);
 
 if (zebra && last > 2) {
   for (let r = 2; r <= last; r += 2) {
     sh.getRange(r, 1, 1, cols).setBackground('#ecfdf5');
   }
 }
}

function defaultHeaders_() {
 return [
   'Email', 'BO Link', 'Registered', 'Country', 'Last Login Country',
   'Account Status', 'Total Deposit Count', 'Total Deposit Sum',
   'Total Cashout Sum', 'GGR', 'Qtag', 'Tags', 'Manager'
 ];
}

function toRow_(p, headers) {
 const reg = p.created_ts ? Utilities.formatDate(new Date(p.created_ts * 1000), Session.getScriptTimeZone(), 'MMMM yyyy') : '';
 
 const map = {
   'Email': p.email || '',
   'BO Link': p.bo || '',
   'Registered': reg,
   'Country': p.country || '',
   'Last Login Country': p.last_login_country || '',
   'Account Status': p.status || '',
   'Total Deposit Count': p.deposit_count || 0,
   'Total Deposit Sum': p.deposit_sum || 0,
   'Total Cashout Sum': p.cashout_sum || 0,
   'GGR': p.ggr || 0,
   'Qtag': p.qtag || '',
   'Tags': p.tags || '',
   'Manager': p.manager || ''
 };
 
 return headers.map(h => (map[h] !== undefined ? map[h] : ''));
}

/* --------------------------- HELPERS ----------------------- */
function MONTHS_() {
 return ['January','February','March','April','May','June','July','August','September','October','November','December'];
}

function normProjectKey_(p) {
 p = String(p || '').toLowerCase();
 return p.indexOf('betsio') >= 0 ? 'betsio' : 'bets';
}

function normSegKey_(s) {
 s = String(s || '').toLowerCase();
 if (s === 'streamer' || s === 'streamers') return 'streamers';
 if (s === 'cross' || s === 'crosssell' || s === 'cross-sell') return 'cross-sell';
 return 'partners';
}

function loadQtagMap_(ss) {
 const map = new Map();
 const sh = ss.getSheetByName('General Info');
 if (!sh) return map;
 
 const rows = sh.getDataRange().getValues();
 for (let i = 1; i < rows.length; i++) {
   const brand = String(rows[i][0] || '').trim().toLowerCase();
   const num = String(rows[i][1] || '').trim();
   const seg = normSegKey_(rows[i][2]);
   
   if (!num) continue;
   
   const brandKey = brand.indexOf('betsio') >= 0 ? 'betsio' : 'bets';
   map.set(brandKey + '|' + num, seg);
 }
 
 return map;
}

function dumpQtagMap_() {
 const map = loadQtagMap_(SpreadsheetApp.getActive());
 const obj = {};
 map.forEach((v, k) => obj[k] = v);
 return obj;
}

function parseQtagNumber_(q) {
 const s = String(q || '');
 const m = s.match(/a(\d+)_/i);
 return m ? m[1] : null;
}

function tokenizeTags_(v) {
 const s = String(v || '').toLowerCase();
 try {
   const arr = JSON.parse(s);
   if (Array.isArray(arr)) return arr.map(x => String(x).toLowerCase());
 } catch (e) {}
 
 const m = s.match(/[a-z0-9_]+/gi);
 return m || [];
}

function buildVipSets_(vipInfo) {
 const vipNameToSet = new Map();
 const slvipNameToSet = new Map();
 const vipTagSet = new Set();
 const slvipTagSet = new Set();
 
 (vipInfo.vip || []).forEach(r => {
   const S = new Set((r.tags || []).map(t => String(t).toLowerCase()));
   vipNameToSet.set(String(r.name || '').toLowerCase(), S);
   (r.tags || []).forEach(t => vipTagSet.add(String(t).toLowerCase()));
 });
 
 (vipInfo.slvip || []).forEach(r => {
   const S = new Set((r.tags || []).map(t => String(t).toLowerCase()));
   slvipNameToSet.set(String(r.name || '').toLowerCase(), S);
   (r.tags || []).forEach(t => slvipTagSet.add(String(t).toLowerCase()));
 });
 
 return { vipNameToSet, slvipNameToSet, vipTagSet, slvipTagSet };
}

function collectIds_(sheet, project, segment, sub, qmap, vipNameToSet, slvipNameToSet, vipTagSet, slvipTagSet, latestActiveMap, qtagNumber) {
 try {
   // Використовуємо кешований результат якщо він є
   const cacheKey = `${sheet.getName()}_${project}_${segment}_${sub}_${qtagNumber || ''}`;
   const cached = PropertiesService.getScriptProperties().getProperty(cacheKey);
   if (cached && cached !== 'null') {
     try {
       const parsed = JSON.parse(cached);
       if (parsed && parsed.timestamp && (Date.now() - parsed.timestamp < 3600000)) { // 1 година
         return new Set(parsed.ids);
       }
     } catch(e) {}
   }
   
   const values = sheet.getDataRange().getValues();
   if (!values.length) return new Set();
   
   const headers = values[0].map(v => String(v).trim().toLowerCase());
   let idIdx = headers.indexOf(ID_HEADER);
   if (idIdx === -1) idIdx = 0;
   
   const tagsIdx = headers.indexOf(TAGS_HEADER);
   const qIdx = headers.indexOf(QTAG_HEADER);
   
   const out = new Set();
   const brandKey = normProjectKey_(project);
   const wantSeg = String(segment || 'general').toLowerCase();
   const wantSub = String(sub || 'all').toLowerCase();
   const qFilter = qtagNumber ? String(qtagNumber) : null;
   
   const vipTarget = (wantSeg === 'vip') ? 
     (wantSub === 'all' ? vipTagSet : (vipNameToSet.get(wantSub) || new Set())) : null;
   const slTarget = (wantSeg === 'slvip') ? 
     (wantSub === 'all' ? slvipTagSet : (slvipNameToSet.get(wantSub) || new Set())) : null;
   
   for (let i = 1; i < values.length; i++) {
     const rawId = String(values[i][idIdx] || '').trim();
     if (!rawId) continue;
     
     const ok = (brandKey === 'bets') ? 
       rawId.toLowerCase().indexOf('bets:') === 0 : 
       rawId.toLowerCase().indexOf('betsio:') === 0;
     if (!ok) continue;
     
     let rawTags = (tagsIdx !== -1) ? values[i][tagsIdx] : '';
     const freshRow = latestActiveMap ? latestActiveMap.get(rawId) : null;
     if (freshRow && freshRow.tags !== undefined && freshRow.tags !== null) {
       rawTags = freshRow.tags;
     }
     
     const tokens = tokenizeTags_(rawTags);
     if (tokens.indexOf('test') >= 0) continue;
     
     const q = (qIdx !== -1) ? values[i][qIdx] : '';
     const n = parseQtagNumber_(q);
     
     if (qFilter && String(n || '') !== qFilter) continue;
     
     if (wantSeg === 'general') {
       out.add(rawId);
       continue;
     }
     
     if (wantSeg === 'organic') {
       if (q === '' || q === null) out.add(rawId);
       continue;
     }
     
     
    if (wantSeg === 'partners') {
       let seg = 'partners';
       if (n) {
         const found = qmap.get(brandKey + '|' + n);
         seg = found || 'partners';
       }
       
       if (wantSub === 'all') {
         // Додаємо ID якщо: колонки qtag немає АБО якщо вона є і не пуста
         if (qIdx === -1 || (q && q !== null && q !== '')) out.add(rawId);
       } else if (seg === wantSub) {
         out.add(rawId);
       }
       continue;
     }
     
     if (wantSeg === 'vip') {
       if (tokens.some(t => vipTarget.has(t))) out.add(rawId);
       continue;
     }
     
     if (wantSeg === 'slvip') {
       if (tokens.some(t => slTarget.has(t))) out.add(rawId);
       continue;
     }
   }
   
   // Кешуємо результат
   try {
     const cacheKey = `${sheet.getName()}_${project}_${segment}_${sub}_${qtagNumber || ''}`;
     const cacheData = {
       ids: Array.from(out),
       timestamp: Date.now()
     };
     PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(cacheData));
   } catch(e) {
     // Ігноруємо помилки кешування
   }
   
   return out;
   
 } catch (e) {
   console.error('collectIds_ error:', e);
   return new Set();
 }
}

function intersectCount_(A, B) {
 let c = 0;
 A.forEach(v => {
   if (B.has(v)) c++;
 });
 return c;
}

function findSheet_(ss, label, kind) {
 const chunk = String(label || '').toLowerCase();
 const sheets = ss.getSheets();
 
 for (const sh of sheets) {
   const low = sh.getName().toLowerCase();
   if (low.indexOf(chunk) >= 0 && low.indexOf(kind) >= 0) return sh;
 }
 
 return null;
}

function mapRowsById_(sheet) {
 try {
   const vals = sheet.getDataRange().getValues();
   const out = new Map();
   
   if (!vals.length) return out;
   
   const headers = vals[0].map(v => String(v).trim().toLowerCase());
   let idIdx = headers.indexOf(ID_HEADER);
   if (idIdx === -1) idIdx = 0;
   
   for (let i = 1; i < vals.length; i++) {
     const id = String(vals[i][idIdx] || '').trim();
     if (!id) continue;
     
     const row = {};
     headers.forEach((h, idx) => row[h] = vals[i][idx]);
     out.set(id, row);
   }
   
   return out;
   
 } catch (e) {
   console.error('mapRowsById_ error:', e);
   return new Map();
 }
}

function collectActiveSheets_(ss) {
 const months = MONTHS_();
 const list = [];
 
 for (const sh of ss.getSheets()) {
   const low = sh.getName().toLowerCase();
   if (low.indexOf('active') === -1) continue;
   
   const idx = months.findIndex(m => low.indexOf(m.toLowerCase()) >= 0);
   if (idx !== -1) {
     list.push({ idx, label: months[idx], sheet: sh });
   }
 }
 
 return list;
}

function aggregateActiveMap_(activeSheetsDesc) {
 const map = new Map();
 
 for (const a of activeSheetsDesc) {
   const m = mapRowsById_(a.sheet);
   m.forEach((row, id) => {
     if (!map.has(id)) map.set(id, row);
   });
 }
 
 return map;
}

function _getAny_(row, keys) {
 if (!row) return '';
 
 const map = Object.keys(row).reduce((m, k) => {
   m[k.toLowerCase()] = k;
   return m;
 }, {});
 
 for (const k of keys) {
   const real = map[String(k).toLowerCase()];
   const v = real !== undefined ? row[real] : undefined;
   if (v !== '' && v !== null && v !== undefined) return v;
 }
 
 return '';
}

function _asNumber_(v, def) {
 v = (v === null || v === undefined || v === '') ? def : v;
 const n = Number(v);
 return isNaN(n) ? (Number(String(v).replace(/[^\d.\-]/g, '')) || def) : n;
}

function _tsFromAny_(v) {
 if (!v && v !== 0) return 0;
 
 const n = Number(v);
 if (!isNaN(n)) return (n > 1e12) ? Math.floor(n / 1000) : n;
 
 const d = new Date(v);
 return isNaN(d.getTime()) ? 0 : Math.floor(d.getTime() / 1000);
}

function _buildBestRow_(preferRow, altRows) {
 const r = {};
 const candidates = [preferRow].concat(altRows || []).filter(Boolean);
 
 r.id = _getAny_(preferRow, ['id']) || _getAny_(candidates[0], ['id']);
 r.tags = _getAny_(candidates.find(x => _getAny_(x, ['tags']) !== '') || {}, ['tags']);
 r.qtag = _getAny_(candidates.find(x => _getAny_(x, ['qtag']) !== '') || {}, ['qtag']);
 
 r.created_at = _getAny_(candidates.find(x => _getAny_(x, ['created_at', 'createdAt', 'registered', 'registration_date', 'created']) !== '') || {}, ['created_at', 'createdAt', 'registered', 'registration_date', 'created']);
 
 r.country = _getAny_(
   candidates.find(x => _getAny_(x, [
     'country', 'registration_country', 'country_code', 'registration_country_code', 'reg_country'
   ]) !== '') || {},
   ['country', 'registration_country', 'country_code', 'registration_country_code', 'reg_country']
 );
 
 r.last_login_country = _getAny_(
   candidates.find(x => _getAny_(x, [
     'last_login_country', 'lastLoginCountry', 'last_login_country_code', 'last_country', 'last_login_country_name'
   ]) !== '') || {},
   ['last_login_country', 'lastLoginCountry', 'last_login_country_code', 'last_country', 'last_login_country_name']
 );
 
 r.disabled = _getAny_(candidates.find(x => _getAny_(x, ['disabled', 'status']) !== '') || {}, ['disabled', 'status']);
 
 r.lifetime_deposit_count_total = _getAny_(candidates.find(x => _getAny_(x, ['lifetime_deposit_count_total', 'lifetime_deposit_count', 'deposit_count_total', 'deposit_count']) !== '') || {}, ['lifetime_deposit_count_total', 'lifetime_deposit_count', 'deposit_count_total', 'deposit_count']);
 
 r.lifetime_deposit_sum_total = _getAny_(candidates.find(x => _getAny_(x, ['lifetime_deposit_sum_total', 'lifetime_deposit_sum', 'deposit_sum_total', 'deposit_sum']) !== '') || {}, ['lifetime_deposit_sum_total', 'lifetime_deposit_sum', 'deposit_sum_total', 'deposit_sum']);
 
 r.lifetime_cashout_sum_total = _getAny_(candidates.find(x => _getAny_(x, ['lifetime_cashout_sum_total', 'lifetime_cashout_sum', 'cashout_sum_total', 'cashout_sum']) !== '') || {}, ['lifetime_cashout_sum_total', 'lifetime_cashout_sum', 'cashout_sum_total', 'cashout_sum']);
 
 r.email = _getAny_(candidates.find(x => _getAny_(x, ['email']) !== '') || {}, ['email']);
 
 return r;
}

function extractRecordFlexible_(row, project) {
 const proj = normProjectKey_(project);
 const id = String(_getAny_(row, ['id']) || '').trim();
 const rest = id.replace(/^betsio:|^bets:/i, '');
 
 const bo = (proj === 'betsio' ? 
   'https://betsio.casino-backend.com/backend/players/' : 
   'https://bets.casino-backend.com/backend/players/') + rest;
 
 const created_ts = _tsFromAny_(_getAny_(row, ['created_at', 'createdAt', 'registered', 'registration_date', 'created']));
 
 const depCnt = _asNumber_(_getAny_(row, ['lifetime_deposit_count_total', 'lifetime_deposit_count', 'deposit_count_total', 'deposit_count']), 0);
 const depSum = _asNumber_(_getAny_(row, ['lifetime_deposit_sum_total', 'lifetime_deposit_sum', 'deposit_sum_total', 'deposit_sum']), 0);
 const coSum = _asNumber_(_getAny_(row, ['lifetime_cashout_sum_total', 'lifetime_cashout_sum', 'cashout_sum_total', 'cashout_sum']), 0);
 const ggr = depSum - coSum;
 
 const disabled = String(_getAny_(row, ['disabled', 'status']) || '').trim().toLowerCase();
 const status = (!disabled || disabled === 'none') ? 'active' : disabled;
 
 return {
   id,
   email: String(_getAny_(row, ['email']) || ''),
   bo,
   created_ts,
   country: String(_getAny_(row, [
     'country', 'registration_country', 'country_code', 'registration_country_code', 'reg_country'
   ]) || ''),
   last_login_country: String(_getAny_(row, [
     'last_login_country', 'lastLoginCountry', 'last_login_country_code', 'last_country', 'last_login_country_name'
   ]) || ''),
   status,
   deposit_count: depCnt,
   deposit_sum: depSum,
   cashout_sum: coSum,
   ggr,
   qtag: String(_getAny_(row, ['qtag']) || ''),
   tags: prettyTags_(_getAny_(row, ['tags']))
 };
}

function prettyTags_(v) {
 if (v === null || v === undefined) return '';
 
 const s = String(v);
 try {
   const arr = JSON.parse(s);
   if (Array.isArray(arr)) return arr.map(x => String(x)).join(', ');
 } catch (e) {}
 
 return s.replace(/[\[\]"]/g, '')
   .split(',')
   .map(x => x.trim())
   .filter(Boolean)
   .join(', ');
}

function detectManager_(tokens, rows) {
 const set = new Set(tokens.map(t => String(t).toLowerCase()));
 
 for (const r of rows || []) {
   for (const t of r.tags || []) {
     if (set.has(String(t).toLowerCase())) return r.name;
   }
 }
 
 return '';
}

function include(name) {
 return HtmlService.createHtmlOutputFromFile(name).getContent();
}
