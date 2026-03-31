// ═══════════════════════════════════════════════════════════════
// ChantierPro — Apps Script v3
// Loué Menuiserie
// ═══════════════════════════════════════════════════════════════

const SHEET_CONTACTS    = 'Contacts';
const SHEET_CHANTIERS   = 'Chantiers';
const SHEET_COTES       = 'Feuille 1';
const SHEET_FABRICATION = 'Fabrication';
const SHEET_MATERIAUX   = 'Materiaux';
const SHEET_FINITIONS   = 'Finitions';
const SHEET_QUINCAILLE  = 'Quincaillerie';
const NOM_ENTREPRISE    = 'Loué Menuiserie';
const DRIVE_FOLDER_NAME = 'ChantierPro — Photos Fabrication';


// ═══════════════════════════════════════════════════════════════
// GET
// ═══════════════════════════════════════════════════════════════
function doGet(e) {
  if (e.parameter.action === 'getContacts') return getContacts(e.parameter.callback);
  return ContentService.createTextOutput('ChantierPro API v3 OK').setMimeType(ContentService.MimeType.TEXT);
}

function getContacts(callback) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sc = ss.getSheetByName(SHEET_CONTACTS);
  const contacts = sc ? sc.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    nom:String(r[0]||''), role:String(r[1]||''), email:String(r[2]||''),
    chefEquipe:String(r[3]||''), destinataire:String(r[4]||''),
    conducteurDefaut:String(r[5]||''), chefAtelier:String(r[6]||''),
  })) : [];
  const sh = ss.getSheetByName(SHEET_CHANTIERS);
  const chantiers = sh ? sh.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    nom:String(r[0]||''), reference:String(r[1]||''),
  })) : [];
  function getList(n){ const s=ss.getSheetByName(n); return s?s.getDataRange().getValues().slice(1).map(r=>String(r[0]||'')).filter(v=>v):[]; }
  const json = JSON.stringify({ contacts, chantiers, materiaux:getList(SHEET_MATERIAUX), finitions:getList(SHEET_FINITIONS), quincaille:getList(SHEET_QUINCAILLE) });
  return callback
    ? ContentService.createTextOutput(callback+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════════
// POST
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'fabrication') traiterFicheFabrication(data);
    else traiterFicheCotes(data);
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    console.error('doPost:', err.message);
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}


// ═══════════════════════════════════════════════════════════════
// FICHE PRISE DE COTES
// ═══════════════════════════════════════════════════════════════
function traiterFicheCotes(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COTES);
  const now = new Date();
  (data.lignes||[]).forEach(l => {
    sheet.appendRow([data.ref,data.chantier,data.batiment,data.auteur,data.date,
      l.logt,l.piece,l.hauteur,l.solType||data.solMode,l.murHaut,l.plintheBas,
      l.long,l.prof,l.pendG,l.pendD,l.plinthe,data.socle,data.commentaires,now]);
  });
  envoyerEmailCotes(data, now);
}

function envoyerEmailCotes(data, horodatage) {
  const contacts = getContactsRaw();
  const row = contacts.find(r=>String(r[5]).toUpperCase()==='O');
  if (!row||!row[2]) return;
  const sujet = `[ChantierPro] Prise de cotes — ${data.ref} — ${data.chantier}`;
  const lignesHtml = (data.lignes||[]).map((l,i)=>`
    <tr style="background:${i%2===0?'#F7F5F1':'white'}">
      <td style="padding:5px 7px;border:1px solid #ddd;font-weight:600">${l.logt||'—'}</td>
      <td style="padding:5px 7px;border:1px solid #ddd">${l.piece||'—'}</td>
      <td style="padding:5px 7px;border:1px solid #ddd;text-align:center;font-weight:700">${l.hauteur||'—'}</td>
      <td style="padding:5px 7px;border:1px solid #ddd;text-align:center;color:#1454A0">${l.solType||data.solMode||'—'}</td>
      <td style="padding:5px 7px;border:1px solid #ddd;text-align:center">${l.murHaut||'—'}</td>
      <td style="padding:5px 7px;border:1px solid #ddd;text-align:center">${l.long||'—'}</td>
      <td style="padding:5px 7px;border:1px solid #ddd;text-align:center">${l.prof||'—'}</td>
      <td style="padding:5px 7px;border:1px solid #ddd;text-align:center">${(l.pendG?'G':'')} ${(l.pendD?'D':'')}</td>
    </tr>`).join('');
  const html = `<html><head><style>
    @page{size:landscape;margin:1cm}body{font-family:Arial,sans-serif;font-size:10px}
    table{width:100%;border-collapse:collapse;margin-top:8px}
    th{background:#1A2B4A;color:white;padding:6px 7px;font-size:9px;text-align:center}
  </style></head><body>
    <table style="width:100%;border-collapse:collapse;margin-bottom:10px;border-bottom:3px solid #1A2B4A;padding-bottom:8px">
      <tr>
        <td><div style="font-weight:800;font-size:14px;color:#1A2B4A">Loué Menuiserie</div><div style="font-size:9px;color:#6A6460">Menuiserie · Charpente · Agencement</div></td>
        <td style="text-align:center"><div style="font-weight:800;font-size:12px;color:#1A2B4A">PRISE DE COTES — FAÇADES &amp; PLACARDS</div>
          <div style="font-size:10px;color:#6A6460;margin-top:3px">Chantier : <b>${data.chantier}</b> &nbsp;|&nbsp; Par : <b>${data.auteur}</b> &nbsp;|&nbsp; Date : <b>${data.date}</b></div></td>
        <td style="text-align:right;font-size:10px;color:#6A6460">Réf : ${data.ref||'—'}</td>
      </tr>
    </table>
    <table><thead>
      <tr><th rowspan="2" style="text-align:left">N° Logt</th><th rowspan="2">Pièce</th>
        <th colspan="6" style="background:#203A6D">FAÇADE</th><th rowspan="2">Pend.</th></tr>
      <tr><th>Haut.</th><th>Sol</th><th>Mur/mur</th><th>Pl/pl</th><th>Long.</th><th>Prof.</th></tr>
    </thead><tbody>${lignesHtml}</tbody></table>
    ${data.commentaires?`<div style="margin-top:10px;padding:8px;border:1px solid #ccc"><b>Commentaires :</b> ${data.commentaires}</div>`:''}
  </body></html>`;
  const blob = HtmlService.createHtmlOutput(html).getAs('application/pdf').setName(`Cotes_${data.chantier}_${data.date||''}.pdf`);
  GmailApp.sendEmail(row[2], sujet, 'Bonjour,\n\nVeuillez trouver ci-joint la fiche de prise de cotes.\n\nChantierPro',{attachments:[blob],name:NOM_ENTREPRISE});
}


// ═══════════════════════════════════════════════════════════════
// FICHE FABRICATION
// ═══════════════════════════════════════════════════════════════
function traiterFicheFabrication(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  let sheet = ss.getSheetByName(SHEET_FABRICATION);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FABRICATION);
    const h = ['N° Série','Chantier','Auteur','Chef Atelier','Conducteur','Date','Délai','Zone','Réf. chantier','Nb Blocs','Commentaires','Horodatage'];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length).setBackground('#1A2B4A').setFontColor('white').setFontWeight('bold');
  }
  sheet.appendRow([data.serie||'',data.chantier,data.auteur,data.chefAtelier||'',data.conducteur||'',
    data.date,data.delai||'',data.zone||'',data.refChantier||'',(data.blocs||[]).length,data.commentaires||'',now]);

  // Photos Drive
  let liensDrive = [];
  let photosBlob = [];  // blobs pour intégration PDF
  if ((data.photosBase64||[]).length > 0) {
    const res = sauvegarderPhotos(data.photosBase64, data.chantier, data.serie||'', now);
    liensDrive = res.liens;
    photosBlob = res.blobs;
  }

  envoyerEmailFabrication(data, liensDrive, photosBlob, now);
}

function sauvegarderPhotos(photosBase64, chantier, serie, now) {
  const liens = [], blobs = [];
  try {
    const existing = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
    const root = existing.hasNext() ? existing.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
    const chClean = chantier.replace(/[^a-zA-Z0-9\s\-]/g,'').trim();
    const existCh = root.getFoldersByName(chClean);
    const chFolder = existCh.hasNext() ? existCh.next() : root.createFolder(chClean);
    const nomDossier = serie
      ? `Série ${serie} — ${Utilities.formatDate(now,'Europe/Paris','dd-MM-yyyy')}`
      : Utilities.formatDate(now,'Europe/Paris','dd-MM-yyyy HH-mm');
    const ficheFolder = chFolder.createFolder(nomDossier);
    photosBase64.forEach((p,i) => {
      try {
        const bytes = Utilities.base64Decode(p.data);
        const b = Utilities.newBlob(bytes, p.mimeType||'image/jpeg', `Photo_${i+1}.jpg`);
        blobs.push({blob:b, label:p.label||`Photo ${i+1}`});
        const file = ficheFolder.createFile(b.copyBlob());
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        liens.push({nom:p.label||`Photo ${i+1}`, url:file.getUrl()});
      } catch(err){ console.error('Photo '+i, err.message); }
    });
  } catch(err){ console.error('Drive:', err.message); }
  return {liens, blobs};
}


// ─── Email fabrication ───
function envoyerEmailFabrication(data, liensDrive, photosBlob, horodatage) {
  const contacts = getContactsRaw();
  const chefRow  = contacts.find(r=>String(r[6]).toUpperCase()==='O');
  const condRow  = contacts.find(r=>String(r[0]).toLowerCase()===(data.conducteur||'').toLowerCase());
  const emailChef = chefRow ? chefRow[2] : null;
  const emailCond = condRow ? condRow[2] : null;
  const destinataires = [...new Set([emailChef,emailCond].filter(Boolean))];
  if (!destinataires.length){ console.log('Aucun destinataire fab'); return; }

  const serie    = data.serie||'';
  const chantier = data.chantier||'—';
  const sujet    = `[ChantierPro] Fiche fabrication — ${serie?'Série '+serie+' — ':''}${chantier}`;

  // ── Cellule N° Série : case vide si non renseigné ──
  const serieCell = serie
    ? `<div style="font-size:13px;font-weight:800;color:#1A2B4A">${serie}</div>`
    : `<div style="width:90px;height:22px;border:1.5px solid #1A2B4A;border-radius:3px;margin-top:2px"></div>`;

  // ── En-tête PDF : 2 colonnes gauche (N°Série+Date / Responsable+Délai) + chantier large + centre titre + droite opérateurs ──
  const enTete = `
  <table style="width:100%;border-collapse:collapse;margin-bottom:10px;border-bottom:3px solid #1A2B4A;padding-bottom:8px">
    <tr>
      <!-- INFOS FICHE : 2 sous-colonnes empilées -->
      <td style="width:340px;vertical-align:top;padding-right:12px">
        <table style="width:100%;border-collapse:collapse;border:1.5px solid #D8D4CE;border-radius:4px">
          <!-- Ligne 1 : N° Série | Date -->
          <tr>
            <td style="padding:5px 8px;border-bottom:1px solid #E1DCD6;border-right:1px solid #E1DCD6;width:50%">
              <div style="font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#6A6460;margin-bottom:2px">N° Série</div>
              ${serieCell}
            </td>
            <td style="padding:5px 8px;border-bottom:1px solid #E1DCD6;width:50%">
              <div style="font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#6A6460;margin-bottom:2px">Date</div>
              <div style="font-size:11px;font-weight:600">${data.date||'—'}</div>
            </td>
          </tr>
          <!-- Ligne 2 : Chantier (pleine largeur) -->
          <tr>
            <td colspan="2" style="padding:5px 8px;border-bottom:1px solid #E1DCD6">
              <div style="font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#6A6460;margin-bottom:2px">Chantier</div>
              <div style="font-size:12px;font-weight:800;color:#1A2B4A;white-space:nowrap;overflow:hidden">${chantier}${data.refChantier?' / '+data.refChantier:''}</div>
            </td>
          </tr>
          <!-- Ligne 3 : Responsable | Délai -->
          <tr>
            <td style="padding:5px 8px;border-right:1px solid #E1DCD6;width:50%">
              <div style="font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#6A6460;margin-bottom:2px">Responsable</div>
              <div style="font-size:11px;font-weight:600">${data.auteur||'—'}</div>
            </td>
            <td style="padding:5px 8px;width:50%">
              <div style="font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#6A6460;margin-bottom:2px">Délai</div>
              <div style="font-size:11px;font-weight:700;color:#9A5600">${data.delai||'—'}</div>
            </td>
          </tr>
          ${data.zone?`<tr><td colspan="2" style="padding:5px 8px">
            <div style="font-size:7.5px;font-weight:700;text-transform:uppercase;color:#6A6460;margin-bottom:2px">Zone</div>
            <div style="font-size:10px">${data.zone}</div>
          </td></tr>`:''}
        </table>
      </td>
      <!-- TITRE -->
      <td style="vertical-align:top;padding:0 12px">
        <div style="font-weight:800;font-size:15px;color:#1A2B4A">Loué Menuiserie</div>
        <div style="font-size:9px;color:#6A6460;margin-bottom:6px">Menuiserie · Charpente · Agencement</div>
        <div style="font-weight:800;font-size:13px;color:#1A2B4A">FICHE DE FABRICATION</div>
        <div style="margin-top:6px;font-size:10px;color:#6A6460">
          Chef d'atelier : <b style="color:#141210">${data.chefAtelier||'—'}</b>
          &nbsp;·&nbsp;
          Conducteur : <b style="color:#141210">${data.conducteur||'—'}</b>
        </div>
      </td>
      <!-- OPÉRATEURS / TEMPS -->
      <td style="width:185px;vertical-align:top">
        <div style="font-size:8px;font-weight:700;text-transform:uppercase;color:#6A6460;margin-bottom:3px">Opérateurs / Temps</div>
        <table style="width:100%;border-collapse:collapse;border:1.5px solid #D8D4CE;font-size:9px">
          <tr>
            <th style="padding:4px 5px;background:#1A2B4A;color:white;text-align:left;font-size:8px">Opérateur</th>
            <th style="padding:4px 5px;background:#1A2B4A;color:white;font-size:8px">Objectif</th>
            <th style="padding:4px 5px;background:#1A2B4A;color:white;font-size:8px">Réalisé</th>
          </tr>
          ${Array(6).fill(null).map(()=>`<tr>
            <td style="padding:5px;border:1px solid #E1DCD6;height:19px"></td>
            <td style="padding:5px;border:1px solid #E1DCD6;height:19px"></td>
            <td style="padding:5px;border:1px solid #E1DCD6;height:19px"></td>
          </tr>`).join('')}
        </table>
      </td>
    </tr>
  </table>`;

  // ── Rendu des blocs ──
  const blocsHtml = (data.blocs||[]).map(b => {
    const label   = b.type==='porte'
      ? ({complet:'Bloc porte complet',vantail:'Vantail seul',huisserie:'Huisserie seule'}[b.porteType]||'Bloc porte')
      : (b.type==='photo'?'Photo / Croquis':'Divers');
    const clr     = b.type==='divers'?'#1A2B4A':b.type==='photo'?'#6B3FA0':'#7A4200';
    const clrBg   = b.type==='divers'?'#EBF1FA':b.type==='photo'?'#EDE8F8':'#FDF0D6';
    let inner = '';

    // DIVERS
    if (b.type==='divers' && (b.rows||[]).length) {
      const trs = b.rows.map((r,i)=>`<tr style="background:${i%2===0?'#F7F5F1':'white'}">
        <td style="padding:4px 6px;border:1px solid #E1DCD6;font-weight:700;font-size:10px">${r.repere||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;text-align:center;font-weight:700;font-size:10px">${r.qte||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;text-align:center;font-size:10px">${r.long||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;text-align:center;font-size:10px">${r.larg||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;text-align:center;font-size:10px">${r.haut||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;text-align:center;font-size:10px">${r.epais||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;text-align:center;font-size:10px">${r.section||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;font-size:10px">${r.materiau||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;font-size:10px">${r.finition||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;font-size:10px">${r.quincaillerie||'—'}</td>
        <td style="padding:4px 6px;border:1px solid #E1DCD6;font-size:9px">${r.observations||''}</td>
      </tr>`).join('');
      inner = `<table style="width:100%;border-collapse:collapse"><thead>
        <tr>
          <th rowspan="2" style="padding:5px 6px;background:#1A2B4A;color:white;font-size:9px;text-align:left">Rep.</th>
          <th rowspan="2" style="padding:5px 6px;background:#1A2B4A;color:white;font-size:9px">Qté</th>
          <th colspan="5" style="padding:5px 6px;background:#203A6D;color:white;font-size:9px">DIMENSIONS (mm)</th>
          <th rowspan="2" style="padding:5px 6px;background:#2A5A3A;color:white;font-size:9px">Matériau</th>
          <th rowspan="2" style="padding:5px 6px;background:#6B3FA0;color:white;font-size:9px">Finition</th>
          <th rowspan="2" style="padding:5px 6px;background:#6B3FA0;color:white;font-size:9px">Quincaillerie</th>
          <th rowspan="2" style="padding:5px 6px;background:#1A2B4A;color:white;font-size:9px">Observations</th>
        </tr>
        <tr>
          <th style="padding:4px 6px;background:#2A4A8A;color:white;font-size:9px">Long.</th>
          <th style="padding:4px 6px;background:#2A4A8A;color:white;font-size:9px">Larg.</th>
          <th style="padding:4px 6px;background:#2A4A8A;color:white;font-size:9px">Haut.</th>
          <th style="padding:4px 6px;background:#2A4A8A;color:white;font-size:9px">Épais.</th>
          <th style="padding:4px 6px;background:#2A4A8A;color:white;font-size:9px">Section</th>
        </tr>
      </thead><tbody>${trs}</tbody></table>`;
    }

    // BLOC PORTE
    if (b.type==='porte' && (b.rows||[]).length) {
      const isV = b.porteType==='vantail'||b.porteType==='complet';
      const isH = b.porteType==='huisserie'||b.porteType==='complet';
      const thV = isV?`<th colspan="6" style="padding:5px;background:#203A6D;color:white;font-size:9px">VANTAIL</th>`:'';
      const thH = isH?`<th colspan="7" style="padding:5px;background:#2A5A3A;color:white;font-size:9px">HUISSERIE</th>`:'';
      const sv = 'padding:4px 5px;background:#2A4A8A;color:white;font-size:8.5px';
      const sh = 'padding:4px 5px;background:#3A7A4A;color:white;font-size:8.5px';
      const th2V = isV?`<th style="${sv}">Haut.</th><th style="${sv}">Larg.</th><th style="${sv}">Sens</th><th style="${sv}">Ferrage</th><th style="${sv}">QT Ferrage</th><th style="${sv}">Nature</th>`:'';
      const th2H = isH?`<th style="${sh}">Support</th><th style="${sh}">Ép.</th><th style="${sh}">Ferrage</th><th style="${sh}">QT Ferrage</th><th style="${sh}">Sens</th><th style="${sh}">Feuillure</th><th style="${sh}">Rainure</th>`:'';
      const c = 'padding:4px 5px;border:1px solid #E1DCD6;font-size:9.5px';
      const trs = b.rows.map((r,i)=>{
        const tdV = isV?`<td style="${c};text-align:center">${r.haut||'—'}</td><td style="${c};text-align:center">${r.larg||'—'}</td><td style="${c}">${r.sens||'—'}</td><td style="${c}">${r.ferrageType||'—'}</td><td style="${c};text-align:center">${r.ferrageQt||'—'}</td><td style="${c}">${r.nature||'—'}</td>`:'';
        const tdH = isH?`<td style="${c}">${r.support||'—'}</td><td style="${c};text-align:center">${r.epSupport||'—'}</td><td style="${c}">${r.ferrage||'—'}</td><td style="${c};text-align:center">${r.ferrageHQt||'—'}</td><td style="${c}">${r.sens||'—'}</td><td style="${c}">${r.feuillure||'—'}</td><td style="${c};text-align:center">${r.rainure||'—'}</td>`:'';
        return `<tr style="background:${i%2===0?'#F7F5F1':'white'}">
          <td style="${c};font-weight:700">${r.repere||'—'}</td>
          <td style="${c};text-align:center;font-weight:700">${r.qte||'—'}</td>
          ${tdV}${tdH}
        </tr>`;
      }).join('');
      inner = `<table style="width:100%;border-collapse:collapse"><thead>
        <tr>
          <th rowspan="2" style="padding:5px;background:#7A4200;color:white;font-size:9px;text-align:left">Rep.</th>
          <th rowspan="2" style="padding:5px;background:#7A4200;color:white;font-size:9px">QT</th>
          ${thV}${thH}
        </tr>
        <tr>${th2V}${th2H}</tr>
      </thead><tbody>${trs}</tbody></table>`;
    }

    // PHOTO — placeholder dans le PDF (photos insérées après en page séparée)
    if (b.type==='photo') {
      inner = `<div style="padding:10px;font-size:10px;color:#6B3FA0;font-style:italic;background:#F7F4FF">
        📷 ${b.nbPhotos||0} photo(s) — voir page(s) suivante(s)</div>`;
    }

    return `<div style="margin-bottom:12px;border:1.5px solid ${clr};border-radius:5px;overflow:hidden;page-break-inside:avoid">
      <div style="background:${clrBg};padding:6px 10px;border-bottom:1px solid ${clr}">
        <span style="font-weight:700;font-size:10px;color:${clr};text-transform:uppercase;letter-spacing:.4px">${label} — Bloc ${b.id}</span>
      </div>
      <div>${inner||'<div style="padding:8px;color:#999;font-size:10px;font-style:italic">Bloc vide</div>'}</div>
    </div>`;
  }).join('');

  // ── Photos intégrées dans le PDF (une par page) ──
  let photosSection = '';
  if (photosBlob.length > 0) {
    photosSection = photosBlob.map((p, i) => {
      try {
        // Encode le blob en base64 pour l'intégrer en data URI dans le HTML
        const bytes  = p.blob.getBytes();
        const base64 = Utilities.base64Encode(bytes);
        const mime   = p.blob.getContentType() || 'image/jpeg';
        return `<div style="page-break-before:always;padding:20px;">
          <div style="font-size:11px;font-weight:700;color:#6B3FA0;margin-bottom:8px;border-bottom:2px solid #6B3FA0;padding-bottom:5px">
            📷 Photo / Croquis — ${p.label}
          </div>
          <img src="data:${mime};base64,${base64}" style="max-width:100%;max-height:700px;display:block;margin:0 auto;border:1px solid #ddd;border-radius:4px">
        </div>`;
      } catch(err){
        console.error('Intégration photo PDF:', err.message);
        return `<div style="page-break-before:always;padding:20px;color:#999;font-style:italic">Photo ${i+1} — erreur d'intégration</div>`;
      }
    }).join('');
  }

  const commentaireSection = data.commentaires
    ? `<div style="margin-top:10px;padding:8px 10px;background:#F7F5F1;border:1px solid #D8D4CE;border-radius:4px;page-break-inside:avoid">
        <div style="font-size:8.5px;font-weight:700;text-transform:uppercase;color:#6A6460;margin-bottom:3px">Commentaires</div>
        <div style="font-size:10px;line-height:1.5">${data.commentaires.replace(/\n/g,'<br>')}</div>
      </div>` : '';

  // ── HTML complet PDF paysage, multi-pages ──
  const htmlPDF = `<html><head><style>
    @page{size:landscape;margin:1cm}
    body{font-family:Arial,sans-serif;font-size:10px;margin:0}
    *{box-sizing:border-box}
    .page-break{page-break-before:always}
  </style></head><body>
    ${enTete}
    ${blocsHtml}
    ${commentaireSection}
    ${photosSection}
  </body></html>`;

  const nomPDF = `Fab_${chantier}${serie?'_S'+serie:''}_${data.date||''}.pdf`;
  const blob   = HtmlService.createHtmlOutput(htmlPDF).getAs('application/pdf').setName(nomPDF);

  // Corps de l'email simple
  const liensDriveHtml = liensDrive.length
    ? `<p><strong>📷 Photos (${liensDrive.length}) :</strong></p><ul>${liensDrive.map(l=>`<li><a href="${l.url}">${l.nom}</a></li>`).join('')}</ul>`
    : '';

  const corps = `<div style="font-family:Arial,sans-serif;font-size:13px">
    <p>Bonjour,</p>
    <p>Veuillez trouver <strong>en pièce jointe</strong> la fiche de fabrication PDF.</p>
    <table style="font-size:12px;border-collapse:collapse;margin:12px 0">
      <tr><td style="color:#6A6460;padding:3px 14px 3px 0">Chantier</td><td><strong>${chantier}</strong></td></tr>
      ${serie?`<tr><td style="color:#6A6460;padding:3px 14px 3px 0">N° Série</td><td><strong>${serie}</strong></td></tr>`:''}
      <tr><td style="color:#6A6460;padding:3px 14px 3px 0">Rédigé par</td><td>${data.auteur||'—'}</td></tr>
      <tr><td style="color:#6A6460;padding:3px 14px 3px 0">Date</td><td>${data.date||'—'}</td></tr>
      <tr><td style="color:#6A6460;padding:3px 14px 3px 0">Délai</td><td><strong style="color:#9A5600">${data.delai||'—'}</strong></td></tr>
    </table>
    ${liensDriveHtml}
    <p style="color:#6A6460;font-size:11px;margin-top:16px">ChantierPro · ${Utilities.formatDate(horodatage,'Europe/Paris','dd/MM/yyyy HH:mm')}</p>
  </div>`;

  destinataires.forEach(email => {
    GmailApp.sendEmail(email, sujet,
      `Bonjour,\n\nVeuillez trouver en pièce jointe la fiche de fabrication — ${chantier}${serie?' — Série '+serie:''}.\n\nChantierPro`,
      {htmlBody:corps, attachments:[blob], name:NOM_ENTREPRISE});
  });
}


// ═══════════════════════════════════════════════════════════════
// UTILITAIRES
// ═══════════════════════════════════════════════════════════════
function getContactsRaw() {
  const sc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONTACTS);
  return sc ? sc.getDataRange().getValues().slice(1).filter(r=>r[0]) : [];
}
