// ============================================================
// JCI Tunisie — Google Apps Script
// Copiez ce code dans Apps Script et redéployez
// ============================================================

const FOLDER_ID = '1vYM5RRsFDCVe4YgmQyR6Cg8jDrV8LUUW';
const SHEET_CONTACT   = 'Contact';
const SHEET_ACTIVITES = 'Activites';

// ── GET : lire les activités ──────────────────────────────
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';

    if (action === 'getActivites') {
      return getActivites();
    }

    // Ping de test
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'JCI Script actif' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST : router les actions ─────────────────────────────
function doPost(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'contact';

    if (action === 'activite')       return addActivite(e);
    if (action === 'deleteActivite') return deleteActivite(e);

    // Par défaut : formulaire de contact
    return addContact(e);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── CONTACT ───────────────────────────────────────────────
function addContact(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_CONTACT);

  // Ajouter les en-têtes si la feuille est vide
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Prénom', 'Nom', 'Email', 'Téléphone', 'Objet', 'Message']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#0A1F44').setFontColor('#F5C518');
  }

  const prenom    = e.parameter.prenom    || '';
  const nom       = e.parameter.nom       || '';
  const email     = e.parameter.email     || '';
  const telephone = e.parameter.telephone || '';
  const objet     = e.parameter.objet     || '';
  const message   = e.parameter.message   || '';

  sheet.appendRow([new Date(), prenom, nom, email, telephone, objet, message]);

  return jsonOk({ message: 'Contact enregistré' });
}

// ── ACTIVITÉS : lire ──────────────────────────────────────
function getActivites() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_ACTIVITES);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return jsonOk([]);

  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  const data = rows
    .filter(r => r[0] !== '')
    .map(r => ({
      id:          r[0].toString(),
      date:        r[1] ? Utilities.formatDate(new Date(r[1]), 'Africa/Tunis', 'dd/MM/yyyy') : '',
      cat:         r[2],
      nom:         r[3],
      desc:        r[4],
      photos:      r[5] ? r[5].toString().split(',').filter(Boolean) : []
    }));

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ACTIVITÉS : ajouter ───────────────────────────────────
function addActivite(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_ACTIVITES);

  // En-têtes si vide
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'Date', 'Catégorie', 'Nom', 'Description', 'Photos']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#0A1F44').setFontColor('#F5C518');
  }

  const id          = Date.now().toString();
  const cat         = e.parameter.cat         || e.parameter.categorie || '';
  const nom         = e.parameter.nom         || '';
  const desc        = e.parameter.desc        || e.parameter.description || '';
  const dateParam   = e.parameter.date        || '';
  const photoLinks  = [];

  // Traiter les photos (fichiers base64 uploadés)
  for (let i = 0; i < 4; i++) {
    const b64      = e.parameter['photo_b64_' + i];
    const mimeType = e.parameter['photo_mime_' + i] || 'image/jpeg';
    const fileName = e.parameter['photo_name_' + i] || ('photo_' + i + '.jpg');

    if (b64) {
      try {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const bytes  = Utilities.base64Decode(b64);
        const blob   = Utilities.newBlob(bytes, mimeType, fileName);
        const file   = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoLinks.push('https://drive.google.com/uc?export=view&id=' + file.getId());
      } catch(err) {
        Logger.log('Erreur upload photo ' + i + ': ' + err.toString());
      }
    }
  }

  sheet.appendRow([
    id,
    dateParam ? new Date(dateParam) : new Date(),
    cat,
    nom,
    desc,
    photoLinks.join(',')
  ]);

  return jsonOk({ id: id, photos: photoLinks });
}

// ── ACTIVITÉS : supprimer ─────────────────────────────────
function deleteActivite(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTIVITES);
  if (!sheet) return jsonOk({ message: 'Feuille introuvable' });

  const id      = e.parameter.id || '';
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOk({ message: 'Rien à supprimer' });

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0].toString() === id) {
      sheet.deleteRow(i + 2);
      return jsonOk({ message: 'Supprimé' });
    }
  }

  return jsonOk({ message: 'ID non trouvé' });
}

// ── HELPERS ───────────────────────────────────────────────
function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}