/**
 * Script Google Apps Script pour l'application "Constitution des classes 2026/2027"
 * Ce script doit être collé dans l'éditeur de script de votre Google Sheet (Extensions > Apps Script).
 * Déployez-le ensuite en tant qu'Application Web ("Web App") avec :
 * - Exécuter en tant que : "Moi (votre adresse email)"
 * - Qui a accès : "Tout le monde" (Anyone)
 */

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "ok", message: "Le script Apps Script fonctionne. Utilisez POST pour envoyer des données." }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // CORS configuration
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    // Si la requête est un preflight OPTIONS
    if (e.parameter && e.parameter.method === "OPTIONS") {
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Récupération des données POST
    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // Première feuille
    
    // Initialiser les en-têtes si la feuille est vide
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Horodatage", 
        "Type", 
        "Niveau", 
        "ID Élèves", 
        "Noms Élèves", 
        "Auteur", 
        "Motif", 
        "Statut"
      ]);
    }
    
    var action = data.action || "add";
    
    if (action === "add") {
      // Validation des données requises
      if (!data.type || !data.niveau || !data.ids || !data.names || !data.author) {
        throw new Error("Champs obligatoires manquants.");
      }
      
      var timestamp = new Date().toISOString();
      
      // Ajout d'une ligne
      sheet.appendRow([
        timestamp,
        data.type,      // "Regroupement" ou "Éloignement"
        data.niveau,    // ex: "3ème"
        data.ids,       // IDs séparés par des ";" (ex: "1626E1;1120E2")
        data.names,     // Noms séparés par des ";" (ex: "ABOUDOU-SARR Eiden; AMADOU-SALL Moussa")
        data.author,    // Initiales/Nom de l'auteur
        data.motif || "",// Explications
        "Actif"         // Statut de la demande
      ]);
      
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        message: "Demande enregistrée avec succès !",
        timestamp: timestamp
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
    } else if (action === "cancel") {
      // Annuler une demande en recherchant son horodatage
      var targetTimestamp = data.timestamp;
      if (!targetTimestamp) {
        throw new Error("Horodatage cible manquant pour l'annulation.");
      }
      
      var lastRow = sheet.getLastRow();
      var range = sheet.getRange(2, 1, lastRow - 1, 8); // Colonnes 1 à 8 (de A à H)
      var values = range.getValues();
      var found = false;
      
      for (var i = 0; i < values.length; i++) {
        // L'horodatage est dans la 1ère colonne (index 0)
        var rowTimestamp = values[i][0];
        
        // Comparaison au format ISO string ou texte brut
        var isMatch = false;
        if (rowTimestamp instanceof Date) {
          isMatch = rowTimestamp.toISOString() === targetTimestamp;
        } else {
          isMatch = String(rowTimestamp) === targetTimestamp;
        }
        
        if (isMatch) {
          // Mettre à jour le statut dans la colonne 8 (H) à "Annulé"
          sheet.getRange(i + 2, 8).setValue("Annulé");
          found = true;
          break;
        }
      }
      
      if (!found) {
        throw new Error("Demande introuvable avec l'horodatage : " + targetTimestamp);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        message: "Demande annulée avec succès !" 
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
    } else {
      throw new Error("Action non reconnue : " + action);
    }
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}
