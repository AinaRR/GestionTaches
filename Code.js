/*************** MENU DÉMARRAGE ***************/
function onOpen() {
  SpreadsheetApp.getUi().createMenu("📋 Menu")
    .addItem("⏳ Synchroniser + Rappels", "syncEtRappels")
    .addItem("📅 Activer rappel automatique", "installerTrigger")
    .addItem("✅ Marquer comme terminé", "marquerCommeTermine")
    .addItem("🕘 Marquer comme en cours", "marquerCommeEnCours")
    .addItem("📝 Marquer comme À faire", "marquerCommeAFaire")
    .addItem("🧹 Réinitialiser les tâches", "resetTaches")
    .addItem("↺  Réinitialiser Historique", "resetHistorique")
    .addToUi();

  creationEntetesTachesSample(); // Création des entêtes dans Tâches sample
  installerTrigger(); // Déclenche automatiquement l'installation du trigger
  syncEtRappels(); 
}

function alignerColonnesADroiteParFeuille(nomFeuille, colonnes) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomFeuille);
  if (!feuille) return;

  const lastRow = feuille.getLastRow();
  if (lastRow < 2) return; // Rien à aligner

  colonnes.forEach(col => {
    feuille.getRange(2, col, lastRow - 1).setHorizontalAlignment("right");
  });
}

function resetHistorique() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historique');
  if (!feuille) {
    SpreadsheetApp.getUi().alert("Feuille 'Historique' introuvable.");
    return;
  }

  const lastRow = feuille.getLastRow();
  if (lastRow > 1) {
    feuille.getRange(2, 1, lastRow - 1, feuille.getLastColumn()).clearContent();
  }

  SpreadsheetApp.getUi().alert("La feuille 'Historique' a été réinitialisée.");
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
    alignerColonnesADroiteParFeuille("Tâches sample", [1, 2, 3, 4, 5, 6, 7]);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const src = ss.getSheetByName('Tâches sample');
    const today = new Date(); 
    today.setHours(0, 0, 0, 0);

    const srcData = src.getDataRange().getValues();
    const headers = [
      "Projet ID", "Projet", "Assigné à", "Email", "Date d’échéance (Projet)", 
      "Statut", "Ligne", "Rappel", "Tâche", "Temps d’échéance (Tâche)"
    ];
    const emails = [];
    const rows = [];

    for (let i = 1; i < srcData.length; i++) {
      const row = srcData[i];
      const [projetIDCell, projet, assigne, email, dateProjet, statut, tache, tempsEcheance] = row;

      const projetID = projetIDCell || "P-" + i.toString().padStart(4, "0"); // 

      if (!projet || !assigne || !email || !dateProjet || !statut) continue;
      if (!/@/.test(email.trim())) continue;

      const parsedDate = new Date(dateProjet);
      if (isNaN(parsedDate.getTime())) continue;
      if (!['À faire', 'En cours', 'Terminé'].includes(statut)) continue;

      const diff = Math.floor((parsedDate - today) / 86400000);
      let rappel = '~';
      let tempsDepasse = false;
      let heureFinale = '';

      if (tempsEcheance instanceof Date) {
        const maintenant = new Date();
        const h = tempsEcheance.getHours();
        const m = tempsEcheance.getMinutes();
        const heureTotale = new Date(maintenant.getTime());
        heureTotale.setHours(h);
        heureTotale.setMinutes(m);
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

      rows.push([projetID, projet, assigne, email, dateProjet, statut, i + 2, rappel, tache, heureFinale]);
    }

    // Envoi des emails
    emails.slice(0, 50).forEach(e => {
      try {
        let message = `Bonjour ${e.assigne},\nVotre tâche “${e.tache}” est prévue pour le ${new Date(e.date).toLocaleDateString()}.`;
        if (e.tempsDepasse) {
          message += `\n⚠️ Attention : le temps d’échéance de cette tâche est déjà dépassé.`;
        }

        MailApp.sendEmail(e.email, `📌 Rappel - ${e.tache}`, message);
      } catch (err) {
        logErreur(`Erreur lors de l'envoi à ${e.email}`, err);
      }
    });

    afficherTableauHTML(headers, rows);

    enregistrerProjetsEtTaches();

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

function creationEntetesTachesSample() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tâches sample');
  if (!feuille) {
    SpreadsheetApp.getUi().alert("Feuille 'Tâches sample' introuvable.");
    return;
  }

  const headers = [
    "ProjetID", // Nouvelle colonne
    "Projet", 
    "Assigné à", 
    "Email", 
    "Date d’échéance (Projet)", 
    "Statut", 
    "Tâche", 
    "Temps d’échéance (Tâche)"
  ];

  feuille.getRange(1, 1, 1, headers.length).setValues([headers]);

  const largeurs = [90, 200, 100, 170, 170, 60, 200, 170];
  for (let i = 0; i < largeurs.length; i++) {
    feuille.setColumnWidth(i + 1, largeurs[i]);
  }

  const totalRows = feuille.getMaxRows();
  feuille.getRange(1, 1, totalRows, headers.length).setWrap(true);

  feuille.getRange(1, 1, 1, headers.length)
    .setFontFamily("Georgia")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold")
    .setBackground("#d6eaf8");
}

function afficherTableauHTML(headers, rows) {
  if (!headers || !Array.isArray(headers)) {
    SpreadsheetApp.getUi().alert("Erreur : les en-têtes sont manquants ou invalides.");
    return;
  }
  if (!rows || !Array.isArray(rows)) {
    SpreadsheetApp.getUi().alert("Erreur : les lignes sont manquantes ou invalides.");
    return;
  }

    // ✅ Formater la colonne date (index 4 car Projet ID est en index 0)
  const timeZone = Session.getScriptTimeZone();
  rows = rows.map(row => {
    const newRow = [...row];
    const dateProjet = row[4];
    if (dateProjet instanceof Date) { newRow[4] = Utilities.formatDate(dateProjet, timeZone, "dd/MM/yyyy");
    }
      return newRow;
    });

  let html = `
    <html>
    <head>
      <style>
        body { font-family: Arial; font-size: 13px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; vertical-align: center; }
        th { background-color: #f0b27a; color: black; cursor: pointer; text-align: center; }
        tr:hover { background-color: #f9f9f9; }
        #searchInput {
          width: 100%;
          padding: 8px;
          border: 1px solid #ccc;
          margin-bottom: 10px;
          font-size: 14px;
        }
      </style>
    </head>
    <body>
      <h2>📋 Tâches enregistrées (HTML)</h2>
      <input type="text" id="searchInput" placeholder="🔍 Rechercher dans le tableau...">

      <table id="tachesTable">
        <thead>
          <tr>${headers.map(h => `<th onclick="sortTable(this)">${h}</th>`).join('')}</tr>
        </thead>
        <tbody>
          ${rows.map(row =>
            `<tr>${row.map(cell => `<td>${cell !== undefined ? cell : ''}</td>`).join('')}</tr>`
          ).join('')}
        </tbody>
      </table>

      <script>
        // Recherche en direct
        document.getElementById('searchInput').addEventListener('keyup', function () {
          const filter = this.value.toLowerCase();
          const rows = document.querySelectorAll('#tachesTable tbody tr');
          rows.forEach(row => {
            const text = row.innerText.toLowerCase();
            row.style.display = text.includes(filter) ? '' : 'none';
          });
        });

        // Tri des colonnes
        function sortTable(th) {
          const table = th.closest('table');
          const tbody = table.querySelector('tbody');
          const index = Array.from(th.parentNode.children).indexOf(th);
          const rows = Array.from(tbody.querySelectorAll('tr'));
          const asc = th.asc = !th.asc;

          rows.sort((a, b) => {
            const cellA = a.children[index].innerText;
            const cellB = b.children[index].innerText;
            return asc
              ? cellA.localeCompare(cellB, undefined, { numeric: true })
              : cellB.localeCompare(cellA, undefined, { numeric: true });
          });

          rows.forEach(row => tbody.appendChild(row));
        }
      </script>
    </body>
    </html>
  `;

  const page = HtmlService.createHtmlOutput(html)
    .setWidth(1200)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(page, 'Tâches générées (HTML interactif)');
}

function verifierOuCreerFeuilleHistorique() {
  const feuilleNom = 'Historique';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let feuille = ss.getSheetByName(feuilleNom);

  if (!feuille) {
    feuille = ss.insertSheet(feuilleNom);
  }

  const headers = [
  "Projet ID",
  "Projet", 
  "Tâche", 
  "Assigné à", 
  "Email", 
  "Date d’échéance (Projet)", 
  "Date et Heure de Création"
];

  // Insérer les en-têtes
  feuille.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Définir des largeurs personnalisées pour les colonnes
  const largeurs = [90, 200, 200, 100, 170, 170, 200];
  for (let i = 0; i < largeurs.length; i++) {
    feuille.setColumnWidth(i + 1, largeurs[i]);
  }

  // Appliquer le retour à la ligne automatique sur toute la feuille
  const totalRows = feuille.getMaxRows();
  feuille.getRange(1, 1, totalRows, headers.length).setWrap(true);


  // Centrer horizontalement et verticalement la ligne d'en-tête (ligne 1)
  feuille.getRange(1, 1, 1, headers.length)
    .setFontFamily("Georgia")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold")
    .setBackground("#F76363");
  
  alignerColonnesADroiteParFeuille("Historique", [1, 2, 3, 4, 5, 6]);
  
  return feuille;
  
}

function enregistrerProjetsEtTaches() {
  const feuilleSource = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tâches sample');
  if (!feuilleSource) return;

  const donneesSource = feuilleSource.getDataRange().getValues();
  if (donneesSource.length < 2) return;

  const feuilleHistorique = verifierOuCreerFeuilleHistorique();
  const donneesHistorique = feuilleHistorique.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const horodatageActuel = Utilities.formatDate(new Date(), timeZone, "dd-MM-yyyy HH:mm");

  // Construction des index pour comparaison
  const projetsSource = {};
  for (let i = 1; i < donneesSource.length; i++) {
    const ligne = donneesSource[i];
    const [projetID, projet, assigneA, email, dateProjet, , tache] = ligne;

    if (!projetID || !projet || !tache || !email || !dateProjet) continue;

    const dateProjetFormatee = dateProjet instanceof Date
      ? Utilities.formatDate(dateProjet, timeZone, "yyyy-MM-dd")
      : dateProjet;

    projetsSource[projetID] = [
      projetID,
      projet,
      tache,
      assigneA,
      email,
      dateProjetFormatee,
      horodatageActuel
    ];
  }

  const projetsHistorique = {};
  for (let i = 1; i < donneesHistorique.length; i++) {
    const ligne = donneesHistorique[i];
    const projetID = ligne[0];
    if (projetID) projetsHistorique[projetID] = i + 1; // ligne réelle
  }

  const lignesASupprimer = [];
  const misesAJour = [];

  // Détecter lignes à supprimer (présentes dans historique mais absentes dans source)
  Object.keys(projetsHistorique).forEach(pid => {
    if (!projetsSource[pid]) {
      lignesASupprimer.push(projetsHistorique[pid]);
    }
  });

  // Appliquer les mises à jour ou ajouts
  Object.entries(projetsSource).forEach(([pid, valeurs]) => {
    if (projetsHistorique[pid]) {
      const ligneIndex = projetsHistorique[pid];
      const ancienneDate = feuilleHistorique.getRange(ligneIndex, 7).getValue();
      valeurs[6] = ancienneDate; // Conserver date de création
      feuilleHistorique.getRange(ligneIndex, 1, 1, valeurs.length).setValues([valeurs]);
    } else {
      feuilleHistorique.appendRow(valeurs);
    }
  });

  // Supprimer les lignes obsolètes (en partant de la fin pour éviter les décalages)
  lignesASupprimer.sort((a, b) => b - a).forEach(index => {
    feuilleHistorique.deleteRow(index);
  });
}