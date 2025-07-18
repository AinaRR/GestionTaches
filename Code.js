/*************** MENU DÉMARRAGE ***************/
function onOpen() {
  SpreadsheetApp.getUi().createMenu("📋 Gestion des tâches")
    .addItem("⏳ Synchroniser + Rappels", "syncEtRappels")
    .addItem("📅 Activer rappel automatique", "installerTrigger")
    .addItem("✅ Marquer comme terminé", "marquerCommeTermine")
    .addItem("🕘 Marquer comme en cours", "marquerCommeEnCours")
    .addItem("📝 Marquer comme À faire", "marquerCommeAFaire")
    .addItem("🧹 Réinitialiser les tâches", "resetTaches")
    .addToUi();

  creationEntetesTachesSample(); // Création des entêtes dans Tâches sample
  creationEntetesTachesEnregistres(); // Création des entête dans Tâches enregistrés
  installerTrigger(); // Déclenche automatiquement l'installation du trigger
}


/*************** MARQUAGE DES STATUTS ***************/
function marquerCommeTermine() {
  mettreAJourStatut("Terminé");
}
function marquerCommeEnCours() {
  mettreAJourStatut("En cours");
}
function marquerCommeAFaire() {
  mettreAJourStatut("À faire");
}

function mettreAJourStatut(nouveauStatut) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = feuille.getActiveRange();
  const colStatut = 5;
  if (!range) return;

  const startRow = range.getRow();
  const numRows = range.getNumRows();

  for (let i = 0; i < numRows; i++) {
    feuille.getRange(startRow + i, colStatut).setValue(nouveauStatut);
  }
}


/*************** SYNCHRONISATION + RAPPELS ***************/
function syncEtRappels() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const src = ss.getSheetByName('Tâches sample');
    const dst = ss.getSheetByName('Tâches enregistrées') || ss.insertSheet('Tâches enregistrées');
    const today = new Date(); 
    today.setHours(0, 0, 0, 0);

    const srcData = src.getDataRange().getValues();
    const headers = [
      "Projet", "Assigné à", "Email", "Date d’échéance (Projet)", 
      "Statut", "Ligne", "Rappel", "Tâche", "Temps d’échéance (Tâche)"
    ];
    const out = [headers];
    const emails = [];

    for (let i = 1; i < srcData.length; i++) {
      const row = srcData[i];
      const [projet, assigne, email, dateProjet, statut, tache, tempsEcheance] = row;

      if (!projet || !assigne || !email || !dateProjet || !statut) continue;
      if (!/@/.test(email.trim())) continue;

      const parsedDate = new Date(dateProjet);
      if (isNaN(parsedDate.getTime())) continue;
      if (!['À faire', 'En cours', 'Terminé'].includes(statut)) continue;

      const dateObj = new Date(dateProjet);
      const diff = Math.floor((dateObj - today) / 86400000);
      let rappel = '~';
      let tempsDepasse = false;
      let heureFinale = '';

      if (tempsEcheance instanceof Date) {
        const maintenant = new Date();
        const h = tempsEcheance.getHours();
        const m = tempsEcheance.getMinutes();
        const heureTotale = new Date(maintenant.getTime());
        heureTotale.setHours(maintenant.getHours() + h);
        heureTotale.setMinutes(maintenant.getMinutes() + m);
        heureFinale = Utilities.formatDate(heureTotale, Session.getScriptTimeZone(), "HH:mm");
      }

      if (statut === 'Terminé') {
        rappel = '✅🔕';
      } else {
        if (diff < 0) {
          rappel = '⌛❌';
        } else if (diff <= 2) {
          rappel = '☑️ à rappeler';
          emails.push({ email: email.trim(), assigne, tache: projet, date: dateProjet, tempsDepasse: false });
        }

        if (tempsEcheance instanceof Date && diff === 0) {
          const maintenant = new Date();
          const heureTache = new Date();
          heureTache.setHours(tempsEcheance.getHours(), tempsEcheance.getMinutes(), 0, 0);

          if (maintenant > heureTache) {
            rappel += ' ⏰ Temps dépassé';
            tempsDepasse = true;
            emails.push({ email: email.trim(), assigne, tache: projet, date: dateProjet, tempsDepasse: true });
          }
        }
      }

      out.push([projet, assigne, email, dateProjet, statut, i + 2, rappel, tache, heureFinale]);
    }

    dst.clearContents();
    dst.getRange(1, 1, out.length, out[0].length).setValues(out);
    dst.getRange(2, 9, out.length - 1).setNumberFormat("hh:mm");

    // 📨 Envoi des e-mails (maximum 50)
    emails.slice(0, 50).forEach(e => {
      try {
        let message = `Bonjour ${e.assigne},\nVotre tâche “${e.tache}” est prévue pour le ${new Date(e.date).toLocaleDateString()}.`;
        if (e.tempsDepasse) {
          message += `\n⚠️ Attention : le temps d’échéance de cette tâche est déjà dépassé.`;
        }

        MailApp.sendEmail(
          e.email,
          `📌 Rappel - ${e.tache}`,
          message
        );
      } catch (err) {
        logErreur(`Erreur lors de l'envoi à ${e.email}`, err);
      }
    });

  } catch (e) {
    logErreur("Erreur dans syncEtRappels()", e);
  }
}


/*************** INSTALLER TRIGGER ***************/
function installerTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'syncEtRappels') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('syncEtRappels')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

}


/*************** RÉINITIALISATION TÂCHES ***************/
function resetTaches() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tâches sample');
  if (sheet) sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
}


/*************** LOGGING D’ERREURS ***************/
function logErreur(msg, e) {
  const message = e?.message || String(e) || 'Erreur inconnue';
  Logger.log(`[ERREUR] ${msg} : ${message}`);
}

function supprimerValidationsEtInfobulles() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const plage = feuille.getRange(1, 1, feuille.getMaxRows(), feuille.getMaxColumns());
  plage.clearDataValidations();
}

function creationEntetesTachesSample() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tâches sample');
  if (!feuille) {
    SpreadsheetApp.getUi().alert("Feuille 'Tâches sample' introuvable.");
    return;
  }

  const headers = [
    "Projet", 
    "Assigné à", 
    "Email", 
    "Date d’échéance (Projet)", 
    "Statut", 
    "Tâche", 
    "Temps d’échéance (Tâche)"
  ];

  // Insérer les en-têtes
  feuille.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Définir des largeurs de colonnes spécifiques
  const largeurs = [200, 100, 170, 170, 60, 200, 170];
  for (let i = 0; i < largeurs.length; i++) {
    feuille.setColumnWidth(i + 1, largeurs[i]); // i + 1 car les colonnes sont 1-based
  }

  const totalRows = feuille.getMaxRows();
  feuille.getRange(1, 1, totalRows, headers.length).setWrap(true);

  feuille.getRange(1, 1, 1, headers.length)
    .setFontFamily("Georgia")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold");  // bonus : mettre en gras les en-têtes

}

function creationEntetesTachesEnregistres() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tâches enregistrées');
  if (!feuille) {
    SpreadsheetApp.getUi().alert("Feuille 'Tâches enregistrées' introuvable.");
    return;
  }

  const headers = [
    "Projet", 
    "Assigné à", 
    "Email", 
    "Date d’échéance (Projet)", 
    "Statut", 
    "Ligne", 
    "Rappel", 
    "Tâche", 
    "Temps d’échéance (Tâche)"
  ];

  // Insérer les en-têtes
  feuille.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Définir les largeurs personnalisées
  const largeurs = [200, 100, 170, 170, 60, 60, 60, 200, 170];
  for (let i = 0; i < largeurs.length; i++) {
    feuille.setColumnWidth(i + 1, largeurs[i]);
  }

  // Appliquer le retour à la ligne automatique sur toute la feuille (colonnes A à I)
  const totalRows = feuille.getMaxRows();
  feuille.getRange(1, 1, totalRows, headers.length).setWrap(true);

  // Centrer horizontalement et verticalement la ligne d'en-tête (ligne 1)
  feuille.getRange(1, 1, 1, headers.length)
    .setFontFamily("Georgia")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold");  // bonus : mettre en gras les en-têtes

}