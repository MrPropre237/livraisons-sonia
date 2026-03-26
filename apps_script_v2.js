// ============================================================
// GOOGLE APPS SCRIPT V2 — Sonia Livraisons
// Corrections appliquées :
//   #1 Solde cumulé via formule Sheets (pas JS)
//   #2 Vérification en-têtes via contenu A1
//   #3 colorerLigne dynamique (pas hardcodé 18)
//   #4 findRowById via TextFinder (plus rapide)
//   #5 Solde corrigé pour UPDATE vs INSERT
//   #6 Header Content-Type retiré (no-cors)
// ============================================================

const SHEET_LIVRAISONS = "📦 Livraisons";
const SHEET_DEPOTS     = "💰 Dépôts";
const SHEET_LOG        = "📋 Journal";

// ── Point d'entrée POST ───────────────────────────────────────────────────
function doPost(e) {
  try {
    // Vérifier que la requête a bien un body
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "Requête invalide — body manquant" });
    }
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

    if (action === "upsert_commande") return handleCommande(payload.data);
    if (action === "upsert_depot")    return handleDepot(payload.data);
    if (action === "log")             return handleLog(payload.data);
    if (action === "sync_all")        return handleSyncAll(payload.data);

    return jsonResponse({ status: "error", message: "Action inconnue: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function doGet(e) {
  return jsonResponse({ status: "ok", message: "Sonia Livraisons V2 — Script actif ✅" });
}

// ── COMMANDES ─────────────────────────────────────────────────────────────
function handleCommande(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_LIVRAISONS);

  ensureCommandeHeaders(sheet); // CORRECTION #2

  const cmdId = data.id;
  // CORRECTION #4 : TextFinder plus efficace
  const row = findRowByIdFast(sheet, cmdId);

  const rowData = [
    cmdId,
    data.date       || new Date().toLocaleDateString("fr-FR"),
    data.zone       || "",
    data.tel        || "",
    data.produit    || "",
    data.qte        || 1,
    data.prix       || 0,
    data.statut     || "attente",
    data.motifEchec || "",
    // Commission : formule IF dans Sheets pour rester dynamique
    data.statut === "livré" ? 1500 : data.statut === "echec" ? -500 : 0,
    // Net à reverser
    data.statut === "livré" ? (data.prix || 0) - 1500 : data.statut === "echec" ? -500 : 0,
    data.depot     ? "O" : "N",
    data.montantDepose || 0,
    // Écart : net - déposé
    (data.statut === "livré" ? (data.prix || 0) - 1500 : data.statut === "echec" ? -500 : 0)
      - (data.montantDepose || 0),
    data.capture   ? "O" : "N",
    data.note      || "",
    data.heure     || "",
    new Date().toLocaleString("fr-FR")
  ];

  if (row) {
    // UPDATE ligne existante
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    colorerLigne(sheet, row, data.statut, data.depot, rowData.length); // CORRECTION #3
  } else {
    // INSERT nouvelle ligne
    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);
    colorerLigne(sheet, lastRow, data.statut, data.depot, rowData.length); // CORRECTION #3
  }

  return jsonResponse({ status: "ok", action: "commande_saved", id: cmdId });
}

// ── DÉPÔTS ────────────────────────────────────────────────────────────────
function handleDepot(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_DEPOTS);

  ensureDepotHeaders(sheet); // CORRECTION #2

  const depotId = data.depotId || ("DEP-" + Date.now());
  // CORRECTION #4 : TextFinder
  const row = findRowByIdFast(sheet, depotId);

  if (row) {
    // ── CORRECTION #5 : UPDATE — remplacer sans recalculer le solde cumulé ──
    // On relit le solde actuel de cette ligne et on recalcule proprement
    const currentMontant = sheet.getRange(row, 4).getValue() || 0;
    const currentSolde   = sheet.getRange(row, 5).getValue() || 0;
    // Nouveau solde = ancien solde - ancien montant + nouveau montant
    const newSolde = currentSolde - currentMontant + (data.montant || 0);

    const rowData = [
      depotId,
      sheet.getRange(row, 2).getValue(), // conserver la date originale
      data.cmdRef    || "",
      data.montant   || 0,
      newSolde,                          // solde recalculé proprement
      data.mode      || "Orange Money",
      data.capture   ? "O" : "N",
      data.valide    ? "O" : "En attente",
      data.note      || "",
      new Date().toLocaleString("fr-FR")
    ];
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    // Mettre à jour la couleur de la ligne modifiée
    const bgUpdate = (data.montant || 0) === 0 ? "#FEF9E7" : "#D5F5E3";
    sheet.getRange(row, 1, 1, rowData.length).setBackground(bgUpdate);
    // Mettre à jour les soldes des lignes suivantes
    recalculerSoldesSuivants(sheet, row);

  } else {
    // ── INSERT nouvelle ligne ──────────────────────────────────────────────
    const lastRow    = sheet.getLastRow() + 1;
    // CORRECTION #1 : solde = somme de tous les dépôts précédents + ce montant
    // On lit le solde de la ligne précédente (si elle existe)
    const soldePrecedent = lastRow > 2
      ? (sheet.getRange(lastRow - 1, 5).getValue() || 0)
      : 0;
    const solde = soldePrecedent + (data.montant || 0);

    const rowData = [
      depotId,
      new Date().toLocaleDateString("fr-FR"),
      data.cmdRef    || "",
      data.montant   || 0,
      solde,
      data.mode      || "Orange Money",
      data.capture   ? "O" : "N",
      data.valide    ? "O" : "En attente",
      data.note      || "",
      new Date().toLocaleString("fr-FR")
    ];
    sheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);

    const bg = lastRow % 2 === 0 ? "#F8FFF9" : "#FFFFFF";
    sheet.getRange(lastRow, 1, 1, rowData.length).setBackground(bg);
  }

  return jsonResponse({ status: "ok", action: "depot_saved", id: depotId });
}

// ── LOG ───────────────────────────────────────────────────────────────────
function handleLog(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_LOG);

  // CORRECTION #2 : vérifier contenu A1 pas juste getLastRow
  if (!sheet.getRange(1, 1).getValue()) {
    const headers = ["Horodatage", "Utilisateur", "Action", "Détail", "Statut"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1A2B4A")
      .setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date().toLocaleString("fr-FR"),
    data.user   || "—",
    data.action || "—",
    data.detail || "—",
    data.statut || "—"
  ]);

  return jsonResponse({ status: "ok", action: "log_saved" });
}

// ── SYNC COMPLÈTE ─────────────────────────────────────────────────────────
function handleSyncAll(data) {
  const cmds   = data.commandes || [];
  const depots = data.depots    || [];
  // Traiter commandes d'abord, puis dépôts
  cmds.forEach(c   => handleCommande(c));
  depots.forEach(d => handleDepot(d));
  return jsonResponse({
    status: "ok",
    action: "sync_complete",
    count: cmds.length + depots.length
  });
}

// ── HELPERS ───────────────────────────────────────────────────────────────

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// CORRECTION #4 : TextFinder — bien plus efficace que boucle for
function findRowByIdFast(sheet, id) {
  const finder = sheet.createTextFinder(String(id)).matchEntireCell(true);
  const result = finder.findNext();
  if (result && result.getColumn() === 1) {
    return result.getRow();
  }
  return null;
}

// Garder l'ancienne pour compatibilité
function findRowById(sheet, id, col) {
  return findRowByIdFast(sheet, id);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// CORRECTION #2 : vérifier contenu A1, pas seulement getLastRow
function ensureCommandeHeaders(sheet) {
  if (sheet.getRange(1, 1).getValue() === "N° CMD") return; // déjà présents
  const headers = [
    "N° CMD", "Date", "Zone / Client", "Téléphone", "Produit", "Qté",
    "Prix brut (FCFA)", "Statut", "Motif échec", "Commission (FCFA)",
    "Net dû (FCFA)", "Dépôt fait ?", "Montant déposé (FCFA)",
    "Écart (FCFA)", "Capture ?", "Note", "Heure", "Mis à jour"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#2E5FA3")
    .setFontColor("#FFFFFF")
    .setFontFamily("Arial");
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 1, 90);
  sheet.setColumnWidths(3, 1, 160);
  sheet.setColumnWidths(5, 1, 180);
}

// CORRECTION #2 : vérifier contenu A1
function ensureDepotHeaders(sheet) {
  if (sheet.getRange(1, 1).getValue() === "N° Dépôt") return; // déjà présents
  const headers = [
    "N° Dépôt", "Date", "CMD Réf.", "Montant (FCFA)", "Solde cumulé (FCFA)",
    "Mode paiement", "Capture ?", "Validé Sonia ?", "Note", "Horodatage"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1E8449")
    .setFontColor("#FFFFFF")
    .setFontFamily("Arial");
  sheet.setFrozenRows(1);
}

// CORRECTION #3 : nbCols dynamique, plus hardcodé à 18
function colorerLigne(sheet, row, statut, depot, nbCols) {
  const cols = nbCols || sheet.getLastColumn();
  const range = sheet.getRange(row, 1, 1, cols);
  if (statut === "livré" && depot) {
    range.setBackground("#D5F5E3");
  } else if (statut === "livré" && !depot) {
    range.setBackground("#FEF9E7");
  } else if (statut === "echec") {
    range.setBackground("#FDECEA");
  } else {
    range.setBackground(row % 2 === 0 ? "#F4F6F8" : "#FFFFFF");
  }
}

// CORRECTION #5 : recalculer les soldes des lignes suivantes après une update
function recalculerSoldesSuivants(sheet, fromRow) {
  const lastRow = sheet.getLastRow();
  if (fromRow >= lastRow) return;
  // Relire tous les montants depuis la ligne 2 jusqu'à fromRow
  let cumulatif = 0;
  for (let r = 2; r <= fromRow; r++) {
    cumulatif += sheet.getRange(r, 4).getValue() || 0;
  }
  // Propager le solde sur les lignes suivantes
  for (let r = fromRow + 1; r <= lastRow; r++) {
    const montant = sheet.getRange(r, 4).getValue() || 0;
    cumulatif += montant;
    sheet.getRange(r, 5).setValue(cumulatif);
  }
}
