/* ==========================================================================
   PARTIE 1 — Lecture du fichier Excel directement dans le navigateur
   ========================================================================== */

const EXCEL_FILE = 'comparatif.xlsx'; // <- renomme toujours ton export vers ce nom
const SHEET_NAME = 'Feuil1';
const FIRST_COL = 2;  // colonne C (0-indexé)
const LAST_COL  = 31; // colonne AF (0-indexé)
const FIRST_BRAND_ROW = 6;  // ligne 7 du fichier (0-indexé)
const LAST_BRAND_ROW  = 46; // ligne 47 du fichier (0-indexé)

function cleanText(v){
  if(v == null) return null;
  return v.toString().replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
}

function slugify(s){
  if(s == null) return null;
  return s.toString().trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function buildDataFromWorkbook(wb){
  const ws = wb.Sheets[SHEET_NAME];
  if(!ws) throw new Error(`Feuille "${SHEET_NAME}" introuvable dans le fichier`);

  function cellValue(r, c){
    const cell = ws[XLSX.utils.encode_cell({r, c})];
    return cell ? cleanText(cell.v) : null;
  }
  function cellHyperlink(r, c){
    const cell = ws[XLSX.utils.encode_cell({r, c})];
    return (cell && cell.l && cell.l.Target) ? cell.l.Target : null;
  }
  const colLetter = c => XLSX.utils.encode_col(c);

  const columns = [];
  for(let c = FIRST_COL; c <= LAST_COL; c++){
    columns.push({
      col: colLetter(c),
      family: cellValue(1, c),
      sub: cellValue(2, c),
      type: cellValue(3, c),
      application: cellValue(5, c)
    });
  }
  let lastFam = null, lastSub = null;
  for(const col of columns){
    if(col.family && col.family !== lastFam) lastSub = null;
    if(col.family) lastFam = col.family; else col.family = lastFam;
    if(col.sub) lastSub = col.sub; else col.sub = lastSub;
  }

  const familiesMap = {};
  const famOrder = [];
  for(const col of columns){
    const fid = slugify(col.family);
    if(!familiesMap[fid]){
      familiesMap[fid] = {id: fid, label: col.family, has_sub: false, subcatsMap: {}, subcatOrder: [], types: [], leaf_col: null, application: null};
      famOrder.push(fid);
    }
    const fam = familiesMap[fid];
    let target;
    if(col.sub){
      fam.has_sub = true;
      const sid = slugify(col.sub);
      if(!fam.subcatsMap[sid]){
        fam.subcatsMap[sid] = {id: sid, label: col.sub, types: []};
        fam.subcatOrder.push(sid);
      }
      target = fam.subcatsMap[sid].types;
    } else {
      target = fam.types;
    }
    if(col.type){
      target.push({id: slugify(col.type), label: col.type, application: col.application, col: col.col});
    } else {
      fam.leaf_col = col.col;
      fam.application = col.application;
    }
  }
  const families = famOrder.map(fid => {
    const f = familiesMap[fid];
    const out = {id: f.id, label: f.label, has_sub: f.has_sub};
    if(f.has_sub) out.subcats = f.subcatOrder.map(sid => f.subcatsMap[sid]);
    else if(f.types.length) out.types = f.types;
    else { out.leaf_col = f.leaf_col; out.application = f.application; }
    return out;
  });

  const brandRows = [];
  const groupLabels = {};
  for(let r = FIRST_BRAND_ROW; r <= LAST_BRAND_ROW; r++){
    const bval = cellValue(r, 1);
    if(bval) brandRows.push({row: r, name: bval});
    const aval = cellValue(r, 0);
    if(aval) groupLabels[r] = aval;
  }
  function groupForRow(r){
    let g = null;
    for(const gr of Object.keys(groupLabels).map(Number).sort((a,b)=>a-b)){
      if(gr <= r) g = groupLabels[gr];
    }
    return g;
  }

  const brands = brandRows.map((b, i) => {
    const startR = b.row;
    const endR = (i + 1 < brandRows.length) ? brandRows[i + 1].row - 1 : LAST_BRAND_ROW;
    const products = {};
    for(let c = FIRST_COL; c <= LAST_COL; c++){
      const entries = [];
      let generic = false;
      for(let r = startR; r <= endR; r++){
        const vs = cellValue(r, c);
        if(!vs) continue;
        const low = vs.toLowerCase();
        if(low === 'non') continue;
        if(low === 'x'){ generic = true; continue; }
        entries.push({name: vs, url: cellHyperlink(r, c)});
      }
      if(entries.length) products[colLetter(c)] = entries;
      else if(generic) products[colLetter(c)] = [{name: 'Disponible (gamme non précisée dans le tableau)', url: null}];
    }
    return {name: b.name, group: groupForRow(startR), products};
  });

  return {families, brands};
}

/* ==========================================================================
   PARTIE 2 — Application (arbre de décision + affichage)
   ========================================================================== */

let DATA = null;

const state = { familyId: null, subcatId: null, typeId: null };

function esc(s){
  const d = document.createElement('div');
  d.textContent = s == null ? '' : s;
  return d.innerHTML;
}

function findFamily(id){ return DATA.families.find(f => f.id === id); }
function findSubcat(fam, id){ return (fam.subcats||[]).find(s => s.id === id); }
function findType(list, id){ return (list||[]).find(t => t.id === id); }

function goBack(){
  if(state.typeId){ state.typeId = null; }
  else if(state.subcatId){ state.subcatId = null; }
  else if(state.familyId){ state.familyId = null; }
  render();
}

function setFamily(id){ state.familyId = id; state.subcatId = null; state.typeId = null; render(); }
function setSubcat(id){ state.subcatId = id; state.typeId = null; render(); }
function setType(id){ state.typeId = id; render(); }

function currentColumn(){
  const fam = findFamily(state.familyId);
  if(!fam) return null;
  if(fam.has_sub){
    const sub = findSubcat(fam, state.subcatId);
    if(!sub) return null;
    const t = findType(sub.types, state.typeId);
    return t ? t.col : null;
  } else if(fam.types && fam.types.length){
    const t = findType(fam.types, state.typeId);
    return t ? t.col : null;
  } else {
    return fam.leaf_col || null;
  }
}

function renderTrail(){
  const trail = document.getElementById('trail');
  const parts = [];
  parts.push({label:'Famille', click: () => { state.familyId=null; state.subcatId=null; state.typeId=null; render(); }, active: !state.familyId});
  const fam = state.familyId ? findFamily(state.familyId) : null;
  if(fam){
    parts.push({label: fam.label, click: () => { state.subcatId=null; state.typeId=null; render(); }, active: fam.has_sub ? !state.subcatId : (fam.types && fam.types.length ? !state.typeId : false)});
  }
  if(fam && fam.has_sub && state.subcatId){
    const sub = findSubcat(fam, state.subcatId);
    if(sub) parts.push({label: sub.label, click: () => { state.typeId=null; render(); }, active: !state.typeId});
  }
  if(state.typeId){
    let list = [];
    if(fam.has_sub){ const sub = findSubcat(fam, state.subcatId); list = sub ? sub.types : []; }
    else { list = fam.types || []; }
    const t = findType(list, state.typeId);
    if(t) parts.push({label: t.label.trim(), click: null, active: true});
  }
  trail.innerHTML = '';
  parts.forEach((p,i) => {
    if(i>0){
      const sep = document.createElement('span');
      sep.className = 'sep'; sep.textContent = '/';
      trail.appendChild(sep);
    }
    const wrap = document.createElement('span');
    wrap.className = 'crumb';
    if(p.click && !p.active){
      const b = document.createElement('button');
      b.textContent = p.label.trim();
      b.onclick = p.click;
      wrap.appendChild(b);
    } else {
      const s = document.createElement('span');
      s.className = 'current';
      s.textContent = p.label.trim();
      wrap.appendChild(s);
    }
    trail.appendChild(wrap);
  });
}

function renderBackButton(){
  const row = document.getElementById('backrow');
  row.style.display = state.familyId ? 'block' : 'none';
}

function stepLabel(n, total){ return `Étape ${n} sur ${total}`; }

function guessTotalSteps(){
  const fam = state.familyId ? findFamily(state.familyId) : null;
  if(!fam) return 3;
  if(fam.has_sub) return 3;
  if(fam.types && fam.types.length) return 2;
  return 1;
}

function render(){
  renderTrail();
  renderBackButton();
  const main = document.getElementById('main');
  main.innerHTML = '';

  if(!state.familyId){
    const h = document.createElement('div');
    h.innerHTML = `<div class="stepline">${stepLabel(1, guessTotalSteps())}</div><h2 class="question">Quel type d'équipement recherchez-vous ?</h2>`;
    main.appendChild(h);
    const grid = document.createElement('div');
    grid.className = 'grid';
    DATA.families.forEach(f => {
      const btn = document.createElement('button');
      btn.className = 'choice';
      btn.innerHTML = `<span class="name">${esc(f.label.replace(/\n/g,' '))}</span>`;
      btn.onclick = () => setFamily(f.id);
      grid.appendChild(btn);
    });
    main.appendChild(grid);
    return;
  }

  const fam = findFamily(state.familyId);

  if(fam.has_sub && !state.subcatId){
    const h = document.createElement('div');
    h.innerHTML = `<div class="stepline">${stepLabel(2, guessTotalSteps())}</div><h2 class="question">${esc(fam.label)} — quelle configuration ?</h2>`;
    main.appendChild(h);
    const grid = document.createElement('div');
    grid.className = 'grid';
    fam.subcats.forEach(s => {
      const btn = document.createElement('button');
      btn.className = 'choice';
      btn.innerHTML = `<span class="name">${esc(s.label)}</span>`;
      btn.onclick = () => setSubcat(s.id);
      grid.appendChild(btn);
    });
    main.appendChild(grid);
    return;
  }

  let typeList = null;
  let stepNum = fam.has_sub ? 3 : 2;
  if(fam.has_sub){
    const sub = findSubcat(fam, state.subcatId);
    typeList = sub ? sub.types : [];
  } else if(fam.types && fam.types.length){
    typeList = fam.types;
    stepNum = 2;
  }

  if(typeList && typeList.length && !state.typeId){
    const h = document.createElement('div');
    const titleBase = fam.has_sub ? findSubcat(fam,state.subcatId).label : fam.label;
    h.innerHTML = `<div class="stepline">${stepLabel(stepNum, guessTotalSteps())}</div><h2 class="question">${esc(titleBase)} — quel type de pompe ?</h2>`;
    main.appendChild(h);
    const grid = document.createElement('div');
    grid.className = 'grid';
    typeList.forEach(t => {
      const btn = document.createElement('button');
      btn.className = 'choice';
      const metaHtml = t.application ? `<span class="meta">${esc(t.application.replace(/\n/g,' · '))}</span>` : '';
      btn.innerHTML = `<span class="name">${esc(t.label.trim())}</span>${metaHtml}`;
      btn.onclick = () => setType(t.id);
      grid.appendChild(btn);
    });
    main.appendChild(grid);
    return;
  }

  renderResults(fam);
}

function renderResults(fam){
  const col = currentColumn();
  const h = document.createElement('div');
  h.innerHTML = `<div class="stepline">Résultat</div><h2 class="question">Fournisseurs disponibles</h2>`;
  document.getElementById('main').appendChild(h);

  const summary = document.createElement('div');
  summary.className = 'result-summary';
  let pathBits = [fam.label.replace(/\n/g,' ')];
  if(fam.has_sub){
    const sub = findSubcat(fam, state.subcatId);
    if(sub) pathBits.push(sub.label);
    const t = findType(sub ? sub.types : [], state.typeId);
    if(t) pathBits.push(t.label.trim());
  } else if(fam.types && fam.types.length){
    const t = findType(fam.types, state.typeId);
    if(t) pathBits.push(t.label.trim());
  }
  summary.innerHTML = `Sélection : <strong>${pathBits.map(esc).join(' → ')}</strong>`;
  if(fam.application){
    summary.innerHTML += `<br>Usage type : ${esc(fam.application)}`;
  }
  document.getElementById('main').appendChild(summary);

  if(!col){
    const empty = document.createElement('div');
    empty.className = 'empty';
    empty.textContent = "Pas d'information disponible pour cette combinaison.";
    document.getElementById('main').appendChild(empty);
    return;
  }

  const accordsBrands = DATA.brands.filter(b => b.products[col] && b.group === 'En accords-cadres');
  const horsBrands = DATA.brands.filter(b => b.products[col] && b.group !== 'En accords-cadres');

  if(accordsBrands.length === 0 && horsBrands.length === 0){
    const empty = document.createElement('div');
    empty.className = 'empty';
    empty.textContent = "Aucun fournisseur référencé pour cette combinaison dans le comparatif.";
    document.getElementById('main').appendChild(empty);
    return;
  }

  function renderGroup(title, brands){
    if(!brands.length) return;
    const wrap = document.createElement('div');
    wrap.className = 'brand-group';
    const label = document.createElement('h3');
    label.className = 'grouplabel';
    label.textContent = title;
    wrap.appendChild(label);
    brands.forEach(b => {
      const card = document.createElement('div');
      card.className = 'brand-card';
      const name = document.createElement('div');
      name.className = 'brand-name';
      name.textContent = b.name;
      card.appendChild(name);
      const list = document.createElement('div');
      list.className = 'product-list';
      b.products[col].forEach(p => {
        const pill = document.createElement('span');
        pill.className = 'product-pill' + (p.url ? ' has-link' : '');
        if(p.url){
          pill.innerHTML = `<a href="${esc(p.url)}" target="_blank" rel="noopener">${esc(p.name)}</a><span class="link-mark">↗</span>`;
        } else {
          pill.textContent = p.name;
        }
        list.appendChild(pill);
      });
      card.appendChild(list);
      wrap.appendChild(card);
    });
    document.getElementById('main').appendChild(wrap);
  }

  renderGroup('En accords-cadres', accordsBrands);
  renderGroup('Hors accords-cadres', horsBrands);
}

document.getElementById('backBtn').addEventListener('click', goBack);

/* --- démarrage : on charge le fichier Excel, pas un JSON --- */
fetch(EXCEL_FILE)
  .then(r => {
    if(!r.ok) throw new Error(`HTTP ${r.status} en cherchant ${EXCEL_FILE}`);
    return r.arrayBuffer();
  })
  .then(buf => {
    const wb = XLSX.read(buf, {type: 'array'});
    DATA = buildDataFromWorkbook(wb);
    render();
  })
  .catch(err => {
    document.getElementById('main').innerHTML =
      `<div class="empty">Erreur de lecture de ${EXCEL_FILE} (${esc(err.message)}).<br>` +
      `Vérifie que le fichier Excel est bien présent dans le dossier du site, sous ce nom exact, ` +
      `et que le site est ouvert via un serveur local (Live Server) et non en double-cliquant sur index.html.</div>`;
  });
