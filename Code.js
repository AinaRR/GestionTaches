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

  formaterColonneHeure();
  ajouterIndentation();

  // ✅ Appliquer les alignements au démarrage
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName("Tâches sample");
  if (src) alignerDonneesSansEntete(src);
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


/*************** FORMATAGE & ALIGNEMENT ***************/
function formaterColonneHeure() { 
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tâches sample");
  if (!feuille) return;

  const colonneHeure = 7;
  const nombreDeLignes = feuille.getLastRow() - 1;
  
  if (nombreDeLignes < 1) return; // 

  feuille.getRange(2, colonneHeure, nombreDeLignes).setNumberFormat("hh:mm");
}

function ajouterIndentation() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tâches sample");
  if (!feuille) return;

  const premiereLigne = 1;
  const nombreDeLignes = feuille.getLastRow() - (premiereLigne - 1);
  if (nombreDeLignes < 1) return;

  const colonnes = [1, 2, 3];
  const indentation = '\u00A0\u00A0';

  colonnes.forEach(col => {
    const plage = feuille.getRange(premiereLigne, col, nombreDeLignes);
    const valeurs = plage.getValues();

    const indentées = valeurs.map(ligne => {
      let valeur = ligne[0];
      if (!valeur || typeof valeur !== 'string') return [valeur];
      if (valeur.startsWith(indentation)) return [valeur];
      return [indentation + valeur];
    });

    plage.setValues(indentées);
  });
}

function alignerDonneesSansEntete(feuille) {
  const nbLignes = feuille.getLastRow() - 1;
  const nbColonnes = feuille.getLastColumn();

  if (nbLignes > 0 && nbColonnes > 0) {
    feuille.getRange(2, 1, nbLignes, nbColonnes).setHorizontalAlignment("right");
    feuille.getRange(1, 1, 1, nbColonnes).setHorizontalAlignment("center");
  }
}


/*************** VALIDATION ********************/
function valider([projet, assigne, email, dateProjet, statut, tache, tempsEcheance]) {
  if (!projet || !assigne || !email || !dateProjet || !statut) return '❌ Champ vide';
  if (!/@/.test(email.trim())) return '❌ Email invalide';
  const parsedDate = new Date(dateProjet);
  if (!(parsedDate instanceof Date) || isNaN(parsedDate.getTime())) return '❌ Date invalide';
  if (!['À faire', 'En cours', 'Terminé'].includes(statut)) return '❌ Statut inconnu';
  return '';
}


/*************** RÉACTION EN DIRECT ***********/
function onEdit(e) {
  try {
    if (!e || !e.range || !e.source) return;

    const ss = e.source;
    const feuilleSource = ss.getSheetByName("Tâches sample");
    const feuilleCible = ss.getSheetByName("Tâches enregistrées");
    if (!feuilleSource || !feuilleCible) return;

    const range = e.range;
    const sheet = range.getSheet();
    const row = range.getRow();
    const col = range.getColumn();

    if (sheet.getName() === "Tâches sample" && col === 7 && row > 1) {
      sheet.getRange(row, col).setNumberFormat("hh:mm");
    }

  } catch (err) {
    Logger.log("[ERREUR] Erreur dans onEdit() : " + err);
  }
}


/*************** SYNCHRONISATION PRINCIPALE ***********/
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

      const projet = row[0];
      const assigne = row[1];
      const email = row[2];
      const dateProjet = row[3];
      const statut = row[4];
      const tache = row[5];
      const tempsEcheance = row[6];

      const erreur = valider([projet, assigne, email, dateProjet, statut, tache, tempsEcheance]);
      if (erreur) continue;

      const dateObj = new Date(dateProjet);
      const diff = Math.floor((dateObj - today) / 86400000);
      let rappel = '~';
      let tempsDepasse = false;

      // 🔄 Calcul de l’heure finale (heure actuelle + durée)
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

        // ⏰ Vérifier si l’heure d’échéance est dépassée aujourd’hui
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

    const nbLignes = out.length - 1;
    if (nbLignes > 0) {
      dst.getRange(2, 9, nbLignes).setNumberFormat("hh:mm");
    }

    const colWidths = [200, 120, 200, 170, 100, 70, 90, 150, 180];
    for (let col = 1; col <= colWidths.length; col++) {
      dst.setColumnWidth(col, colWidths[col - 1]);
    }

    alignerDonneesSansEntete(dst);
    alignerDonneesSansEntete(src);

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


/*************** INSTALLER TRIGGER ********************/
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

  SpreadsheetApp.getUi().alert("📅 Rappel automatique activé à 9h chaque jour");
}


/*************** RÉINITIALISATION TÂCHES ***************/
function resetTaches() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tâches sample');
  if (sheet) sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
}


/*************** LOGGING D’ERREURS *********************/
function logErreur(msg, e) {
  Logger.log(`[ERREUR] ${msg} : ${e.message}`);
}