const TZ = "Europe/Paris";
const TRAILING_HEADER = "ID_TECH";

const SITES = {
  ENGHIEN: {
    label: "Enghien",
    calendarCourses: "5070eef4750ee7ed49ddd64b8e24aac7761883b68cd80794aa51d6ec55ba9878@group.calendar.google.com",
    calendarEvents: "c_95097a55e1f1bc75252d106b8eff97d5d633d51ff3ea226c22ccce28e1e16398@group.calendar.google.com",
    sheetName: "Planning Enghien",
    logSheetName: "LOG_SYNC_ENGHIEN",
    dashboardSheetName: "DASHBOARD_ENGHIEN",
    alertsSheetName: "ALERTES_ENGHIEN"
  },

  VINCENNES: {
    label: "Vincennes",
    calendarCourses: "c_191e503e203f48b3dd60830afb021a8c409a00ee78e74bf3eddfd1ac3c8b4fb3@group.calendar.google.com",
    calendarEvents: "c_be54864e0dc5165b2ee44d05c026d76b516bb95229801ec54afd8a14d542d89a@group.calendar.google.com",
    sheetName: "Planning Vincennes",
    logSheetName: "LOG_SYNC_VINCENNES",
    dashboardSheetName: "DASHBOARD_VINCENNES",
    alertsSheetName: "ALERTES_VINCENNES"
  }
};

const GROUPE_1_REFERENCE = [
  { course: "Prix de Cornulier", hippodrome: "Vincennes", date2026: "2026-01-18" },
  { course: "Prix d’Amérique Legend Race", hippodrome: "Vincennes", date2026: "2026-01-25" },
  { course: "Prix de l’Île-de-France", hippodrome: "Vincennes", date2026: "2026-02-01" },
  { course: "Prix de France – Speed Race", hippodrome: "Vincennes", date2026: "2026-02-08" },
  { course: "Critérium des Jeunes (Prix Comte P. de Montesson)", hippodrome: "Vincennes", date2026: "2026-02-15" },
  { course: "Prix de Paris – Marathon Race", hippodrome: "Vincennes", date2026: "2026-02-22" },
  { course: "Prix de Sélection", hippodrome: "Vincennes", date2026: "" },
  { course: "Prix des Centaures", hippodrome: "Vincennes", date2026: "" },
  { course: "Prix Henri Desmontils", hippodrome: "Vincennes", date2026: "" },
  { course: "Prix de l’Atlantique", hippodrome: "Enghien", date2026: "" },
  { course: "Prix d’Essai (Étrier 3 ans finale)", hippodrome: "Vincennes", date2026: "2026-06-21" },
  { course: "Prix du Président de la République", hippodrome: "Vincennes", date2026: "2026-06-21" },
  { course: "Prix de Normandie", hippodrome: "Vincennes", date2026: "2026-06-21" },
  { course: "Prix Albert Viel", hippodrome: "Vincennes", date2026: "2026-06-21" },
  { course: "Critérium des 3 ans", hippodrome: "Vincennes", date2026: "2026-09-12" },
  { course: "Critérium des 4 ans", hippodrome: "Vincennes", date2026: "2026-09-12" },
  { course: "Critérium des 5 ans", hippodrome: "Vincennes", date2026: "2026-09-12" },
  { course: "Prix des Élites", hippodrome: "Vincennes", date2026: "" },
  { course: "Prix Ready Cash", hippodrome: "Vincennes", date2026: "" },
  { course: "Prix de Vincennes", hippodrome: "Vincennes", date2026: "" },
  { course: "Critérium Continental", hippodrome: "Vincennes", date2026: "" }
];

// =======================
// MENU
// =======================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Planning Hippodromes")
    .addItem("Synchroniser Enghien", "syncEnghien")
    .addItem("Synchroniser Vincennes", "syncVincennes")
    .addItem("Synchroniser tout", "syncAllSites")
     
    .addSeparator()
    .addItem("Masquer les lignes passées", "masquerLignesPasseesTousSites_")
    .addItem("Afficher toutes les lignes", "afficherToutesLesLignesTousSites_")
    .addSeparator()
    .addItem("Tester tous les agendas", "testerTousLesAgendas_")
        .addSeparator()

    .addItem("➕ Ajouter un événement", "openEventForm")
        .addSeparator()

    .addItem(" 📩 Générer Email recap  Vincennes  ", "ouvrirRecapVincennesP1")
    .addToUi();

  masquerLignesPasseesTousSites_();
}
 
 /************************************************************
 * RÉCAP VINCENNES MODERNE - VERSION COMPLÈTE
 *
 * IMPORTANT
 * 1) Active le service avancé Google Calendar API dans Apps Script
 * 2) Si tu as déjà un onOpen() dans ton projet, ajoute simplement :
 *      ajouterMenuRecapVincennesP1_();
 *    à l’intérieur de ton onOpen existant.
 ************************************************************/
 

function ouvrirRecapVincennesP1() {
  const html = HtmlService
    .createHtmlOutputFromFile("getHtmlDialogP1_")
    .setWidth(1180)
    .setHeight(900);

  SpreadsheetApp.getUi().showModalDialog(html, "Récap Vincennes");
}

function getDialogDefaultsP1_() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");

  return {
    dateDebut: `${yyyy}-${mm}-${dd}`,
    dateFin: "",
    emails: [
      { label: "Jad Zoghbi", value: "jad.zoghbi@letrot.com" },
      { label: "Paul Lerosier", value: "ext.paul.lerosier@letrot.com" },
      { label: "Email personnalisé", value: "__custom__" }
    ],
    emailDefaut: "jad.zoghbi@letrot.com"
  };
}


function parserDateISO_P1_(texte, finDeJournee) {
  if (!texte) return null;
  const m = texte.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10) - 1;
  const d = parseInt(m[3], 10);
  const date = new Date(y, mo, d);
  if (date.getFullYear() !== y || date.getMonth() !== mo || date.getDate() !== d) return null;
  if (finDeJournee) date.setHours(23, 59, 59, 999); else date.setHours(0, 0, 0, 0);
  return date;
}

function nettoyerTitreOption(titre) {
  return String(titre || "")
    .replace(/\[OPTION\]/gi, "")
    .replace(/\(OPTION\)/gi, "")
    .replace(/\bOPTION\b/gi, "")
    .replace(/SOUS OPTION/gi, "")
    .replace(/\s+-\s+-/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}


function genererApercuRecapVincennesP1(formData) {
  const start = parserDateISO_P1_(formData.dateDebut, false);
  if (!start) throw new Error("Merci de saisir une date de début valide.");

  let destinataire = (formData.emailChoisi || "").trim();
  if (destinataire === "__custom__") destinataire = (formData.emailLibre || "").trim();
  if (!destinataire || destinataire.indexOf("@") === -1) throw new Error("Adresse email invalide.");

  const mode = String(formData.mode || "manuel");
  if (mode === "25_prochains_evenements") return construireRecapVincennesP1_25Events_(start, destinataire);

  const end = parserDateISO_P1_(formData.dateFin, true);
  if (!end) throw new Error("Merci de saisir une date de fin valide.");
  if (end < start) throw new Error("La date de fin doit être postérieure ou égale à la date de début.");

  return construireRecapVincennesP1_(start, end, destinataire);
}

function envoyerRecapVincennesP1(formData) {
  const payload = genererApercuRecapVincennesP1(formData);
  MailApp.sendEmail({to: payload.destinataire, subject: payload.subject, htmlBody: payload.htmlBody});
  return "Récap envoyé à " + payload.destinataire;
}

/************************************************************
 * BUILDERS
 ************************************************************/
function recapHelpersVincennesP1_() {
  const joursFR = ["DIMANCHE","LUNDI","MARDI","MERCREDI","JEUDI","VENDREDI","SAMEDI"];
  const joursCourtFR = ["Dim","Lun","Mar","Mer","Jeu","Ven","Sam"];
  const moisFR = ["janvier","février","mars","avril","mai","juin","juillet","août","septembre","octobre","novembre","décembre"];
  function formaterDateLongue(date) { const j = date.getDate(); return (j === 1 ? "1er" : j) + " " + moisFR[date.getMonth()]; }
  function formaterHeure(date) { return `${
    date.getHours()}h${String(date.getMinutes()).padStart(2,"0")}`; }
  function formaterDateHeureCellule(date) { return `${joursFR[date.getDay()]} ${date.getDate()}/${date.getMonth()+1}/${date.getFullYear()}<br>&gt; ${formaterHeure(date)}`; }
  function getKeyDate(date) { return date.getFullYear() + "-" + (date.getMonth()+1) + "-" + date.getDate(); }
function detecterType(ev) {
  const titre = String(ev.summary || "").toUpperCase();
  const descNettoyee = nettoyerDescriptionCalendar_(ev.description || "");
  const descUpper = descNettoyee.toUpperCase();

  // Cas réel dans tes descriptions :
  // "Événement B2B" ou "Événement B2C"
  if (/\b[ÉE]V[ÉE]NEMENT\s+B2B\b/i.test(descNettoyee)) {
    return "B2B";
  }

  if (/\b[ÉE]V[ÉE]NEMENT\s+B2C\b/i.test(descNettoyee)) {
    return "B2C";
  }

  // Sécurité si un autre format apparaît un jour
  let matchType = descUpper.match(/\bTYPE\s*:?\s*(B2B|B2C)\b/);
  if (matchType && matchType[1]) {
    return matchType[1].toUpperCase();
  }

  // Fallback métier sur le titre
  if (
    titre.includes("TOURNAGE") ||
    titre.includes("ENTREPRISE") ||
    titre.includes("SEMINAIRE") ||
    titre.includes("SÉMINAIRE") ||
    titre.includes("COCKTAIL") ||
    titre.includes("RECEPTIF") ||
    titre.includes("RÉCEPTIF")
  ) {
    return "B2B";
  }

  return "Non renseigné";
}
function nettoyerDescriptionCalendar_(desc) {
  return String(desc || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<p[^>]*>/gi, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/<[^>]+>/g, "")
    .replace(/\r/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n[ \t]+/g, "\n")
    .trim();
}
function couleurFondType(type) { if (type === "B2B") return "#f3e8ff"; if (type === "B2C") return "#e0f2fe"; return "#ffffff"; }
function couleurPointType(type) { if (type === "B2B") return "#a855f7"; if (type === "B2C") return "#0284c7"; return "#111827"; }
function calculerDuree(ev) {
    const s = new Date(ev.start.date || ev.start.dateTime);
    const e = new Date(ev.end.date || ev.end.dateTime);
    const ms = e - s;
    const jours = Math.round(ms / (1000 * 60 * 60 * 24));
    if (jours > 1) return jours + " jours";
    if (jours === 1) return "1 jour";
    const h = Math.floor(ms / (1000 * 60 * 60));
    const m = Math.floor((ms % (1000 * 60 * 60)) / (1000 * 60));
    return h + "h" + (m > 0 ? String(m).padStart(2, "0") : "00");
  }
  function extraireNomEspace(nomComplet) { let t = (nomComplet || "").replace(/\(.*?\)/g, "").trim(); const parties = t.split("-"); return parties[parties.length - 1].trim(); }
  function extraireResponsable(desc) { let responsable = "Non renseigné"; const match = (desc || "").match(/responsables?(?:\s*setf)?\s*:?\s*(.+)/i); if (match && match[1]) responsable = match[1].split("\n")[0].replace(/<\/?[^>]+(>|$)/g, "").trim(); return responsable; }
  function extraireChamp(desc, libelle) {
    const txt = String(desc || "");
    const regex = new RegExp("(?:^|\\n)\\s*(?:" + libelle + ")\\s*:?\\s*(.+)", "i");
    const match = txt.match(regex);
    if (match && match[1]) return match[1].split("\n")[0].replace(/<\/?[^>]+(>|$)/g, "").trim();
    return "";
  }
  function extraireOuiNon(desc, libelle) { const valeur = extraireChamp(desc, libelle).toUpperCase(); if (valeur.indexOf("OUI") > -1) return "✓"; if (valeur.indexOf("NON") > -1) return ""; return ""; }
  function eventToucheJourDeCourse(ev, joursCourses) {
    const s = new Date(ev.start.date || ev.start.dateTime);
    const e = new Date(ev.end.date || ev.end.dateTime);
    const sd = new Date(s.getFullYear(), s.getMonth(), s.getDate());
    const ed = new Date(e.getFullYear(), e.getMonth(), e.getDate());
    for (let d = new Date(sd); d <= ed; d.setDate(d.getDate() + 1)) if (joursCourses.has(getKeyDate(d))) return true;
    return false;
  }
  function regrouperEvenementsParTitre(events) {
    const groupes = {};
    events.forEach(ev => { const titre = (ev.summary || "").trim(); if (!titre) return; if (!groupes[titre]) groupes[titre] = []; groupes[titre].push(ev); });
    return groupes;
  }
  return {joursFR, joursCourtFR, moisFR, formaterDateLongue, formaterHeure, formaterDateHeureCellule, getKeyDate, detecterType, couleurFondType, couleurPointType, calculerDuree, extraireNomEspace, extraireResponsable, extraireChamp, extraireOuiNon, eventToucheJourDeCourse, regrouperEvenementsParTitre};
}
 const RECAP_VINCENNES_CFG_P1 = {
  calendarEvents: "c_be54864e0dc5165b2ee44d05c026d76b516bb95229801ec54afd8a14d542d89a@group.calendar.google.com",
  calendarCourses: "c_191e503e203f48b3dd60830afb021a8c409a00ee78e74bf3eddfd1ac3c8b4fb3@group.calendar.google.com",
  timezone: "Europe/Paris",
  emails: [
    { label: "Jad Zoghbi", value: "jad.zoghbi@letrot.com" },
    { label: "Paul Lerosier", value: "ext.paul.lerosier@letrot.com" },
    { label: "Email personnalisé", value: "__custom__" }
  ],
  emailDefaut: "jad.zoghbi@letrot.com"
};

function ouvrirRecapVincennesP1() {
  const html = HtmlService
    .createHtmlOutputFromFile("getHtmlDialogP1_")
    .setWidth(1180)
    .setHeight(900);

  SpreadsheetApp.getUi().showModalDialog(html, "Récap Vincennes");
}

function getDialogDefaultsP1_() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");

  return {
    dateDebut: `${yyyy}-${mm}-${dd}`,
    dateFin: "",
    emails: RECAP_VINCENNES_CFG_P1.emails,
    emailDefaut: RECAP_VINCENNES_CFG_P1.emailDefaut
  };
}

function parserDateISO_P1_(texte, finDeJournee) {
  if (!texte) return null;
  const m = String(texte).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;

  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10) - 1;
  const d = parseInt(m[3], 10);
  const date = new Date(y, mo, d);

  if (date.getFullYear() !== y || date.getMonth() !== mo || date.getDate() !== d) {
    return null;
  }

  if (finDeJournee) date.setHours(23, 59, 59, 999);
  else date.setHours(0, 0, 0, 0);

  return date;
}

function genererApercuRecapVincennesP1(formData) {
  const start = parserDateISO_P1_(formData.dateDebut, false);
  if (!start) throw new Error("Merci de saisir une date de début valide.");

  let destinataire = String(formData.emailChoisi || "").trim();
  if (destinataire === "__custom__") {
    destinataire = String(formData.emailLibre || "").trim();
  }
  if (!destinataire || destinataire.indexOf("@") === -1) {
    throw new Error("Adresse email invalide.");
  }

  const mode = String(formData.mode || "manuel");
  if (mode === "25_prochains_evenements") {
    return construireRecapVincennesP1_25Events_(start, destinataire);
  }

  const end = parserDateISO_P1_(formData.dateFin, true);
  if (!end) throw new Error("Merci de saisir une date de fin valide.");
  if (end < start) throw new Error("La date de fin doit être postérieure ou égale à la date de début.");

  return construireRecapVincennesP1_(start, end, destinataire);
}

function envoyerRecapVincennesP1(formData) {
  const payload = genererApercuRecapVincennesP1(formData);
  MailApp.sendEmail({
    to: payload.destinataire,
    subject: payload.subject,
    htmlBody: payload.htmlBody
  });
  return "Récap envoyé à " + payload.destinataire;
}

function recapHelpersVincennesP1_() {
  const joursFR = ["DIMANCHE","LUNDI","MARDI","MERCREDI","JEUDI","VENDREDI","SAMEDI"];
  const moisFR = ["janvier","février","mars","avril","mai","juin","juillet","août","septembre","octobre","novembre","décembre"];

  function formaterDateLongue(date) {
    const j = date.getDate();
    return (j === 1 ? "1er" : j) + " " + moisFR[date.getMonth()] + " " + date.getFullYear();
  }

  function formaterHeure(date) {
    return `${date.getHours()}h${String(date.getMinutes()).padStart(2, "0")}`;
  }

  function formaterDateHeureCellule(date) {
    return `${joursFR[date.getDay()]} ${date.getDate()}/${date.getMonth()+1}/${date.getFullYear()}<br>&gt; ${formaterHeure(date)}`;
  }

  function getKeyDate(date) {
    return date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate();
  }

  function nettoyerDescriptionCalendar(desc) {
    return String(desc || "")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>/gi, "\n")
      .replace(/<p[^>]*>/gi, "")
      .replace(/&nbsp;/gi, " ")
      .replace(/<[^>]+>/g, "")
      .replace(/\r/g, "")
      .replace(/[ \t]+/g, " ")
      .replace(/\n[ \t]+/g, "\n")
      .trim();
  }

  function detecterType(ev) {
    const titre = String(ev.summary || "").toUpperCase();
    const descNettoyee = nettoyerDescriptionCalendar(ev.description || "");
    const descUpper = descNettoyee.toUpperCase();

    if (/\b[ÉE]V[ÉE]NEMENT\s+B2B\b/i.test(descNettoyee)) return "B2B";
    if (/\b[ÉE]V[ÉE]NEMENT\s+B2C\b/i.test(descNettoyee)) return "B2C";

    const matchType = descUpper.match(/\bTYPE\s*:?\s*(B2B|B2C)\b/);
    if (matchType && matchType[1]) return matchType[1].toUpperCase();

    if (
      titre.includes("TOURNAGE") ||
      titre.includes("ENTREPRISE") ||
      titre.includes("SEMINAIRE") ||
      titre.includes("SÉMINAIRE") ||
      titre.includes("COCKTAIL") ||
      titre.includes("RECEPTIF") ||
      titre.includes("RÉCEPTIF")
    ) {
      return "B2B";
    }

    return "Non renseigné";
  }

  function couleurBarreType(type) {
    if (type === "B2B") return "#8b5cf6";
    if (type === "B2C") return "#0ea5e9";
    return "#111827";
  }

  function couleurTexteType(type) {
    if (type === "B2B") return "#7c3aed";
    if (type === "B2C") return "#0284c7";
    return "#111827";
  }

  function extraireNomEspace(nomComplet) {
    let t = String(nomComplet || "").replace(/\(.*?\)/g, "").trim();
    const parties = t.split("-");
    return parties[parties.length - 1].trim();
  }

  function extraireResponsable(desc) {
    const clean = nettoyerDescriptionCalendar(desc);
    const patterns = [
      /responsables?(?:\s*setf)?\s*:?\s*(.+)/i,
      /responsable\s*:?\s*(.+)/i
    ];

    for (var i = 0; i < patterns.length; i++) {
      const match = clean.match(patterns[i]);
      if (match && match[1]) {
        return match[1].split("\n")[0].trim();
      }
    }

    return "Non renseigné";
  }

  function extraireChamp(desc, libelle) {
    const txt = nettoyerDescriptionCalendar(desc);
    const regex = new RegExp("(?:^|\\n)\\s*(?:" + libelle + ")\\s*:?\\s*(.+)", "i");
    const match = txt.match(regex);
    if (match && match[1]) return match[1].split("\n")[0].trim();
    return "";
  }

  function extraireOuiNon(desc, libelle) {
    const valeur = extraireChamp(desc, libelle).toUpperCase();
    if (valeur.includes("OUI")) return "✓";
    if (valeur.includes("NON")) return "";
    return "";
  }

  function eventToucheJourDeCourse(ev, joursCourses) {
    const s = new Date(ev.start.date || ev.start.dateTime);
    const e = new Date(ev.end.date || ev.end.dateTime);
    const sd = new Date(s.getFullYear(), s.getMonth(), s.getDate());
    const ed = new Date(e.getFullYear(), e.getMonth(), e.getDate());

    for (let d = new Date(sd); d <= ed; d.setDate(d.getDate() + 1)) {
      if (joursCourses.has(getKeyDate(d))) return true;
    }
    return false;
  }

  return {
    joursFR,
    moisFR,
    formaterDateLongue,
    formaterHeure,
    formaterDateHeureCellule,
    getKeyDate,
    nettoyerDescriptionCalendar,
    detecterType,
    couleurBarreType,
    couleurTexteType,
    extraireNomEspace,
    extraireResponsable,
    extraireChamp,
    extraireOuiNon,
    eventToucheJourDeCourse
  };
}

function regrouperEventsGlobal_(events) {
  const groupes = {};

  events.forEach(ev => {
    const nom = String(ev.summary || "").trim();
    if (!nom) return;

    const start = new Date(ev.start.date || ev.start.dateTime);
    const end = new Date(ev.end.date || ev.end.dateTime);

    if (!groupes[nom]) {
      groupes[nom] = {
        nom: nom,
        events: [],
        debutMin: start,
        finMax: end,
        espaces: new Set()
      };
    }

    groupes[nom].events.push(ev);
    if (start < groupes[nom].debutMin) groupes[nom].debutMin = start;
    if (end > groupes[nom].finMax) groupes[nom].finMax = end;
  });

  return Object.values(groupes);
}

function logDebug_(scope, message) {
  const ligne = "[" + scope + "] " + message;
  Logger.log(ligne);
  if (typeof console !== "undefined" && console.log) console.log(ligne);
}

function regrouperOccupationsEspaceParEvent_(items) {
  const groupes = {};
  let typesManquants = 0;

  items.forEach(item => {
    const key = String(item.event || "").trim();
    if (!key) return;
    if (!item.type) typesManquants += 1;

    if (!groupes[key]) {
      groupes[key] = {
        event: item.event,
        debut: item.debut,
        fin: item.fin,
        course: !!item.course,
        type: item.type || "",
      };
    } else {
      if (item.debut < groupes[key].debut) groupes[key].debut = item.debut;
      if (item.fin > groupes[key].fin) groupes[key].fin = item.fin;
      if (item.course) groupes[key].course = true;
      if (!groupes[key].type && item.type) groupes[key].type = item.type;
    }
  });

  const resultat = Object.values(groupes).sort((a, b) => a.debut - b.debut);
  logDebug_(
    "regrouperOccupationsEspaceParEvent_",
    "items=" + items.length + ", groupes=" + resultat.length + ", typesManquants=" + typesManquants
  );

  return resultat;
}

function calculerDureeDepuisDates_(debut, fin) {
  const ms = fin - debut;
  if (ms < 0) {
    logDebug_(
      "calculerDureeDepuisDates_",
      "durée négative détectée entre " + debut + " et " + fin
    );
  }
  const minutesTotal = Math.round(ms / (1000 * 60));
  const jours = Math.floor(minutesTotal / (60 * 24));

  if (jours > 1) return jours + " jours";
  if (jours === 1) return "1 jour";

  const heures = Math.floor(minutesTotal / 60);
  const minutes = minutesTotal % 60;
  return heures + "h" + String(minutes).padStart(2, "0");
}

function listerEventsCalendar_(calendarId, start, end, maxResults) {
  return Calendar.Events.list(calendarId, {
    timeMin: start.toISOString(),
    timeMax: end.toISOString(),
    singleEvents: true,
    orderBy: "startTime",
    maxResults: maxResults || 2500
  }).items || [];
}

function construireSetJoursCourses_(eventsCourses, H) {
  const joursCourses = new Set();
  let plagesInvalides = 0;

  eventsCourses.forEach(ev => {
    const s = new Date(ev.start.date || ev.start.dateTime);
    const e = new Date(ev.end.date || ev.end.dateTime);
    const sd = new Date(s.getFullYear(), s.getMonth(), s.getDate());
    const ed = new Date(e.getFullYear(), e.getMonth(), e.getDate());

    if (ed < sd) {
      plagesInvalides += 1;
      logDebug_(
        "construireSetJoursCourses_",
        "plage ignorée pour " + (ev.summary || "(sans titre)") + " car end < start"
      );
      return;
    }

    for (let d = new Date(sd); d <= ed; d.setDate(d.getDate() + 1)) {
      joursCourses.add(H.getKeyDate(d));
    }
  });

  logDebug_(
    "construireSetJoursCourses_",
    "eventsCourses=" + eventsCourses.length + ", joursCourses=" + joursCourses.size + ", plagesInvalides=" + plagesInvalides
  );

  return joursCourses;
}
function detecterPhaseEvent(titre) {
  const t = String(titre || "").toUpperCase();

  if (t.includes("MONTAGE")) return "MONTAGE";
  if (t.includes("DEMONTAGE") || t.includes("DÉMONTAGE")) return "DEMONTAGE";

  return "EXPLOITATION";
}

function detecterOptionEvent(titre) {
  const t = String(titre || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

  return (
    t.includes("[OPTION]") ||
    t.includes("(OPTION)") ||
    t.includes(" OPTION ") ||
    t.startsWith("OPTION ") ||
    t.endsWith(" OPTION") ||
    t.includes("OPTIONNEL") ||
    t.includes("SOUS OPTION")
  );
}

function construireRecapVincennesP1_(start, end, destinataire) {
  const H = recapHelpersVincennesP1_();
  const dateDebutObjet = H.formaterDateLongue(start);
  const dateFinObjet = H.formaterDateLongue(end);

  const events = listerEventsCalendar_(RECAP_VINCENNES_CFG_P1.calendarEvents, start, end, 2500);
  const subject = `Récapitulatif B2C B2B - du ${dateDebutObjet} au ${dateFinObjet} - Hippodrome Paris-Vincennes`;

  if (!events.length) {
    return {
      destinataire,
      subject,
      htmlBody: `<div style="font-family:Arial,sans-serif;padding:24px;"><p>Bonjour,</p><p>Aucun événement trouvé pour la période du <b>${dateDebutObjet}</b> au <b>${dateFinObjet}</b>.</p></div>`
    };
  }

  events.sort((a, b) => new Date(a.start.date || a.start.dateTime) - new Date(b.start.date || b.start.dateTime));

  const eventsCourses = listerEventsCalendar_(RECAP_VINCENNES_CFG_P1.calendarCourses, start, end, 2500);
  const joursCourses = construireSetJoursCourses_(eventsCourses, H);
  const eventsGroupes = regrouperEventsGlobal_(events);
 
 let html = `
  <div style="background:#f6f8fb;padding:30px;font-family:Arial,sans-serif;color:#1f2937;">
    <div style="max-width:1450px;margin:auto;">

      <div style="
        background:linear-gradient(135deg,#ffffff 0%,#fffaf5 100%);
        border-radius:22px;
        padding:26px 28px;
        box-shadow:0 10px 30px rgba(15,23,42,0.06);
        border:1px solid #f1f5f9;
      ">
        <div style="
          display:inline-block;
          padding:7px 12px;
          border-radius:999px;
          background:#fff7ed;
          color:#c2410c;
          font-size:12px;
          font-weight:800;
          letter-spacing:0.2px;
          margin-bottom:14px;
        ">
          Hippodrome Paris-Vincennes
        </div>

        <h1 style="
          margin:0;
          font-size:30px;
          line-height:1.1;
          color:#111827;
          letter-spacing:-0.4px;
        ">
          Récapitulatif B2B/B2C
        </h1>

        <p style="
          margin:12px 0 0 0;
          color:#6b7280;
          font-size:15px;
          line-height:1.5;
        ">
            du <b style="color:#111827;">${dateDebutObjet}</b> au <b style="color:#111827;">${dateFinObjet}</b>
        </p>
      </div>
`;

  let currentMonth = new Date(start.getFullYear(), start.getMonth(), 1);
  const endMonth = new Date(end.getFullYear(), end.getMonth(), 1);

  while (currentMonth <= endMonth) {
    const moisIndex = currentMonth.getMonth();
    const annee = currentMonth.getFullYear();
    const nbJours = new Date(annee, moisIndex + 1, 0).getDate();

    const groupesDuMois = eventsGroupes.filter(groupe => {
      const firstDayOfMonth = new Date(annee, moisIndex, 1);
      const lastDayOfMonth = new Date(annee, moisIndex, nbJours);
      const startDay = new Date(groupe.debutMin.getFullYear(), groupe.debutMin.getMonth(), groupe.debutMin.getDate());
      const endDay = new Date(groupe.finMax.getFullYear(), groupe.finMax.getMonth(), groupe.finMax.getDate());
      return !(endDay < firstDayOfMonth || startDay > lastDayOfMonth);
    });

    html += `<div style="margin-top:22px;background:#ffffff;border-radius:18px;padding:20px;box-shadow:0 2px 10px rgba(15,23,42,0.05);border:1px solid #e5e7eb;">`;
    html += `<h3 style="margin:0 0 14px 0;font-size:24px;">${H.moisFR[moisIndex].toUpperCase()} ${annee}</h3>`;

    if (!groupesDuMois.length) {
      html += `<p style="margin:0;color:#6b7280;font-style:italic;">Aucun événement prévu sur cette période.</p></div>`;
      currentMonth.setMonth(currentMonth.getMonth() + 1);
      continue;
    }

    html += `<div style="overflow:hidden;"><table style="width:100%;table-layout:fixed;border-collapse:separate;border-spacing:0;font-size:10px;border:1px solid #e5e7eb;border-radius:14px;overflow:hidden;background:#ffffff;">`;
    html += `<tr><th style="background:#f8fafc;border-bottom:1px solid #e5e7eb;padding:6px 8px;text-align:left;width:26%;font-size:10px;line-height:1.1;">Event</th>`;

    for (let j = 1; j <= nbJours; j++) {
      const dateCol = new Date(annee, moisIndex, j);
      const joursAbr = ["Dim","Lun","Mar","Mer","Jeu","Ven","Sam"];
      const jourCourt = joursAbr[dateCol.getDay()];
      const isCourseDay = joursCourses.has(H.getKeyDate(dateCol));

      const thStyle = `background:${isCourseDay ? "#fff7ed" : "#f8fafc"};border-bottom:1px solid #e5e7eb;text-align:center;font-weight:700;font-size:9px;padding:4px 2px;${isCourseDay ? "box-shadow: inset 0 -2px 0 #f59e0b;" : ""}`;

      html += `<th style="${thStyle}"><div style="font-size:9px;line-height:1;">${jourCourt}</div><div style="font-size:10px;line-height:1.1;margin-top:1px;">${j}</div><div style="font-size:8px;line-height:1;margin-top:3px;color:${isCourseDay ? "#d97706" : "transparent"};">${isCourseDay ? "🐎" : "&nbsp;"}</div></th>`;
    }
    html += `</tr>`;

groupesDuMois
  .sort((a, b) => {
    const diff = a.debutMin - b.debutMin;
    if (diff !== 0) return diff;
    return a.nom.localeCompare(b.nom, "fr", { sensitivity: "base" });
  })
  .forEach(groupe => {
    const type = H.detecterType(groupe.events[0]);
    const phase = detecterPhaseEvent(groupe.nom);
const isOption = detecterOptionEvent(groupe.nom);
let nomAffiche = groupe.nom;

if (isOption) {
  nomAffiche = "[OPTION] - " + nettoyerTitreOption(groupe.nom);
}
 
 
    let couleurTexte = H.couleurTexteType(type);
    let couleurBarre = H.couleurBarreType(type);
    let hauteurBarre = "8px";

    if (isOption) {
     if (isOption) {
  // ✅ on garde la couleur texte d’origine
  couleurBarre = (type === "B2B")
    ? "rgba(124,58,237,0.16)"
    : "rgba(2,132,199,0.16)";
  hauteurBarre = "6px";
}
    } else if (phase === "MONTAGE") {
      couleurBarre = (type === "B2B")
        ? "rgba(139,92,246,0.40)"
        : "rgba(14,165,233,0.40)";
    } else if (phase === "DEMONTAGE") {
      couleurBarre = (type === "B2B")
        ? "rgba(139,92,246,0.24)"
        : "rgba(14,165,233,0.24)";
    }

    const sd = new Date(
      groupe.debutMin.getFullYear(),
      groupe.debutMin.getMonth(),
      groupe.debutMin.getDate()
    );
    const ed = new Date(
      groupe.finMax.getFullYear(),
      groupe.finMax.getMonth(),
      groupe.finMax.getDate()
    );

    html += `<tr><td style="
      border-bottom:1px solid #e5e7eb;
      padding:6px 6px;
      font-weight:700;
      font-size:10px;
      line-height:1.2;
      color:${couleurTexte};
      vertical-align:top;
      ${isOption ? "font-style:italic; opacity:0.9;" : ""}
    ">${isOption ? "◇ " : ""}${nomAffiche}</td>`;

    for (let j = 1; j <= nbJours; j++) {
      const current = new Date(annee, moisIndex, j);
      const prev = new Date(annee, moisIndex, j - 1);
      const next = new Date(annee, moisIndex, j + 1);
      const isCourseDay = joursCourses.has(H.getKeyDate(current));

      const cellStyle = `border-bottom:1px solid #eef2f7;text-align:center;padding:1px 0;position:relative;${isCourseDay ? "box-shadow: inset 0 -2px 0 #f59e0b;" : ""}`;

      const actif = current >= sd && current <= ed;
      const actifAvant = prev >= sd && prev <= ed;
      const actifApres = next >= sd && next <= ed;

      if (actif) {
        let radius = "0";
        if (!actifAvant && !actifApres) radius = "999px";
        else if (!actifAvant && actifApres) radius = "999px 0 0 999px";
        else if (actifAvant && !actifApres) radius = "0 999px 999px 0";

        let ligneOptionHtml = "";

        if (isOption) {
          const couleurPointille = type === "B2B" ? "#a78bfa" : "#7dd3fc";
          ligneOptionHtml = `<div style="
            position:absolute;
            top:50%;
            left:0;
            width:100%;
            border-top:2px dashed ${couleurPointille};
            transform:translateY(-50%);
          "></div>`;
        }

        html += `<td style="${cellStyle}">
          <div style="
            position:relative;
            margin:0;
            height:${hauteurBarre};
            width:100%;
            background:${couleurBarre};
            border-radius:${radius};
            overflow:hidden;
          ">
            ${ligneOptionHtml}
          </div>
        </td>`;
      } else {
        html += `<td style="${cellStyle}"></td>`;
      }
    }

    html += `</tr>`;
  });

    html += `</table></div></div>`;
    currentMonth.setMonth(currentMonth.getMonth() + 1);
  }

  let recapEspaces = {};

  html += `
    <div style="margin-top:24px;background:#ffffff;border-radius:18px;padding:20px;box-shadow:0 2px 10px rgba(15,23,42,0.05);border:1px solid #e5e7eb;">
      <h2 style="margin:0 0 16px 0;font-size:24px;">Détail des événements</h2>
      <div style="overflow-x:auto;">
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <tr style="background:#f8fafc;">
            <th style="padding:10px;border:1px solid #e5e7eb;">Date début</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Date fin</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Événement</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Type</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Client</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Pax</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Réunion de courses</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Durée</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Espaces</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">ONET</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">CJ Sécurité</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Technique</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Autres prestataires</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Responsable</th>
            <th style="padding:10px;border:1px solid #e5e7eb;">Commentaire</th>
          </tr>`;

eventsGroupes.forEach(groupe => {
  const phase = detecterPhaseEvent(groupe.nom);
  const isOption = detecterOptionEvent(groupe.nom);
    const evRef = groupe.events[0];
    const evStart = groupe.debutMin;
    const evEnd = groupe.finMax;
    const desc = evRef.description || "";
    const type = H.detecterType(evRef);
    const client = H.extraireChamp(desc, "Client");
    const pax = H.extraireChamp(desc, "Pax");
    const responsable = H.extraireResponsable(desc);
    const reunion = groupe.events.some(ev => H.eventToucheJourDeCourse(ev, joursCourses)) ? "✓" : "";
    const onet = H.extraireOuiNon(desc, "ONET|Prestation Nettoyage");
    const cj = H.extraireOuiNon(desc, "CJ SECURITE|CJ SÉCURITÉ|Prestation Sécurité");
    const technique = H.extraireOuiNon(desc, "Technique|Prestation Electricité");
    const autresPrestataires = H.extraireOuiNon(desc, "Autres prestataires|Raccordage Réseau");
    const commentaire = H.extraireChamp(desc, "Commentaire");

    groupe.events.forEach(ev => {
      if (ev.attendees) {
        ev.attendees.forEach(a => {
          if (a.resource === true) {
            const nom = H.extraireNomEspace(a.displayName || a.email);
            groupe.espaces.add(nom);
            if (!recapEspaces[nom]) recapEspaces[nom] = [];
            recapEspaces[nom].push({
              event: groupe.nom,
              debut: groupe.debutMin,
              fin: groupe.finMax,
              course: reunion === "✓"
            });
          }
        });
      }
    });

    const couleurTexte = H.couleurTexteType(type);

    html += `
      <tr style="background:#ffffff;">
        <td style="padding:10px;border:1px solid #e5e7eb;">${H.formaterDateHeureCellule(evStart)}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;">${H.formaterDateHeureCellule(evEnd)}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;font-weight:700;color:${couleurTexte};">${groupe.nom}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;text-align:center;font-weight:700;color:${couleurTexte};">${type === "Non renseigné" ? "" : type}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;">${client}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;text-align:center;">${pax}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;text-align:center;">${reunion}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;">${calculerDureeDepuisDates_(evStart, evEnd)}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;">${Array.from(groupe.espaces).join("<br>")}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;text-align:center;">${onet}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;text-align:center;">${cj}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;text-align:center;">${technique}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;text-align:center;">${autresPrestataires}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;">${responsable}</td>
        <td style="padding:10px;border:1px solid #e5e7eb;">${commentaire}</td>
      </tr>`;
  });

  html += `</table></div></div>`;

  html += `
    <div style="margin-top:24px;background:#ffffff;border-radius:18px;padding:20px;box-shadow:0 2px 10px rgba(15,23,42,0.05);border:1px solid #e5e7eb;">
      <h2 style="margin:0 0 8px 0;font-size:24px;">Récapitulatif des occupations par espace</h2>
      <p style="margin:0 0 16px 0;font-size:12px;color:#1e3a8a;font-weight:700;">Les occurrences identiques sont regroupées par événement.</p>`;

Object.keys(recapEspaces)
  .sort((a, b) => a.localeCompare(b, "fr", { sensitivity: "base" }))
  .forEach(espace => {
    const itemsGroupes = regrouperOccupationsEspaceParEvent_(recapEspaces[espace]);

    html += `
      <div style="
        margin:18px 0 22px 0;
        border:1px solid #e5e7eb;
        border-radius:16px;
        overflow:hidden;
        background:#ffffff;
      ">
        <div style="
          padding:14px 18px;
          background:#f8fafc;
          border-bottom:1px solid #e5e7eb;
          font-size:16px;
          font-weight:800;
          color:#111827;
          text-transform:uppercase;
          letter-spacing:0.3px;
        ">
          ${espace}
        </div>
        <div style="padding:8px 14px 10px 14px;">
    `;

    itemsGroupes.forEach(item => {
      const isOption = detecterOptionEvent(item.event);
      const phase = detecterPhaseEvent(item.event);

      const memeJour =
        item.debut.getFullYear() === item.fin.getFullYear() &&
        item.debut.getMonth() === item.fin.getMonth() &&
        item.debut.getDate() === item.fin.getDate();

      const estJourneeEntiere =
        item.debut.getHours() === 0 &&
        item.debut.getMinutes() === 0 &&
        item.fin.getHours() === 0 &&
        item.fin.getMinutes() === 0;

      let textePeriode = "";

      if (memeJour && !estJourneeEntiere) {
        textePeriode = `le ${H.formaterDateLongue(item.debut)} de ${H.formaterHeure(item.debut)} à ${H.formaterHeure(item.fin)}`;
      } else if (memeJour && estJourneeEntiere) {
        textePeriode = `le ${H.formaterDateLongue(item.debut)}`;
      } else {
        textePeriode = `du ${H.formaterDateLongue(item.debut)} au ${H.formaterDateLongue(item.fin)}`;
      }

      let fond = "#ffffff";
      let bordure = "#e5e7eb";
      let badgeFond = "#f3f4f6";
      let badgeCouleur = "#4b5563";
      let titreCouleur = "#111827";

      if (item.course) {
        fond = "#fff7ed";
        bordure = "#fdba74";
        badgeFond = "#ffedd5";
        badgeCouleur = "#c2410c";
        titreCouleur = "#9a3412";
      }

      if (phase === "MONTAGE" && !item.course) {
        fond = "#f5f3ff";
        bordure = "#ddd6fe";
        titreCouleur = "#6d28d9";
      }

      if (phase === "DEMONTAGE" && !item.course) {
        fond = "#faf5ff";
        bordure = "#e9d5ff";
        titreCouleur = "#7e22ce";
      }

      if (isOption && !item.course) {
        fond = "#ffffff";
        bordure = "#a1a1aa";
        titreCouleur = "#52525b";
        badgeFond = "#f4f4f5";
        badgeCouleur = "#71717a";
      }

      html += `
        <div style="
          display:flex;
          align-items:flex-start;
          gap:12px;
          padding:10px 12px;
          margin:8px 0;
          border:1px ${isOption ? "dashed" : "solid"} ${bordure};
          border-radius:12px;
          background:${fond};
        ">

          <div style="
            flex:0 0 auto;
            padding:6px 10px;
            border-radius:999px;
            background:${badgeFond};
            color:${badgeCouleur};
            font-size:12px;
            font-weight:700;
            line-height:1.2;
            max-width:320px;
          ">
            ${item.course ? "🐎 " : ""}${textePeriode}
          </div>

          <div style="
            flex:1 1 auto;
            font-size:15px;
            font-weight:700;
            color:${titreCouleur};
            line-height:1.3;
            padding-top:4px;
            ${isOption ? "font-style:italic; opacity:0.85;" : ""}
          ">
${isOption ? "◇ " : ""}${nettoyerTitreOption(item.event)}
            ${isOption ? `
              <span style="
                margin-left:8px;
                padding:2px 6px;
                border-radius:999px;
                border:1px dashed #a1a1aa;
                font-size:10px;
                font-weight:700;
                color:#71717a;
                background:#fafafa;
              ">Option</span>
            ` : ""}
          </div>

        </div>
      `;
    });

    html += `
        </div>
      </div>
    `;
  });

  return { destinataire, subject, htmlBody: html };
}

function construireRecapVincennesP1_25Events_(start, destinataire) {
  const H = recapHelpersVincennesP1_();
  const endSearch = new Date(start.getTime());
  endSearch.setMonth(endSearch.getMonth() + 12);

  const events = listerEventsCalendar_(RECAP_VINCENNES_CFG_P1.calendarEvents, start, endSearch, 250).slice(0, 25);
  const eventsCourses = listerEventsCalendar_(RECAP_VINCENNES_CFG_P1.calendarCourses, start, endSearch, 250);
  const joursCourses = construireSetJoursCourses_(eventsCourses, H);
  const subject = "Récapitulatif des 25 prochains événements - Hippodrome Paris-Vincennes";

  let html = `<div style="background:#f6f8fb;padding:30px;font-family:Arial,sans-serif;color:#1f2937;"><div style="max-width:1200px;margin:auto;"><div style="background:#ffffff;border-radius:18px;padding:24px;box-shadow:0 2px 10px rgba(15,23,42,0.05);border:1px solid #e5e7eb;"><h1 style="margin:0;font-size:28px;">25 prochains événements</h1><p style="margin:10px 0 0 0;color:#6b7280;font-size:15px;">À partir du <b>${H.formaterDateLongue(start)}</b></p></div><div style="margin-top:24px;background:#ffffff;border-radius:18px;padding:20px;box-shadow:0 2px 10px rgba(15,23,42,0.05);border:1px solid #e5e7eb;"><table style="width:100%;border-collapse:collapse;font-size:13px;"><tr style="background:#f8fafc;"><th style="padding:10px;border:1px solid #e5e7eb;">Date début</th><th style="padding:10px;border:1px solid #e5e7eb;">Date fin</th><th style="padding:10px;border:1px solid #e5e7eb;">Événement</th><th style="padding:10px;border:1px solid #e5e7eb;">Type</th><th style="padding:10px;border:1px solid #e5e7eb;">Client</th><th style="padding:10px;border:1px solid #e5e7eb;">Pax</th><th style="padding:10px;border:1px solid #e5e7eb;">Course</th><th style="padding:10px;border:1px solid #e5e7eb;">Durée</th><th style="padding:10px;border:1px solid #e5e7eb;">Espaces</th><th style="padding:10px;border:1px solid #e5e7eb;">Responsable</th></tr>`;

  events.forEach(ev => {
    const s = new Date(ev.start.date || ev.start.dateTime);
    const e = new Date(ev.end.date || ev.end.dateTime);
    const desc = ev.description || "";
    const type = H.detecterType(ev);
    const client = H.extraireChamp(desc, "Client");
    const pax = H.extraireChamp(desc, "Pax");
    const reunion = H.eventToucheJourDeCourse(ev, joursCourses) ? "✓" : "";
    const responsable = H.extraireResponsable(desc);
    const couleurTexte = H.couleurTexteType(type);

    const espaces = [];
    if (ev.attendees) {
      ev.attendees.forEach(a => {
        if (a.resource === true) {
          const nom = H.extraireNomEspace(a.displayName || a.email);
          if (espaces.indexOf(nom) === -1) espaces.push(nom);
        }
      });
    }

    html += `<tr style="background:#ffffff;"><td style="padding:10px;border:1px solid #e5e7eb;">${H.formaterDateHeureCellule(s)}</td><td style="padding:10px;border:1px solid #e5e7eb;">${H.formaterDateHeureCellule(e)}</td><td style="padding:10px;border:1px solid #e5e7eb;font-weight:700;color:${couleurTexte};">${ev.summary}</td><td style="padding:10px;border:1px solid #e5e7eb;text-align:center;font-weight:700;color:${couleurTexte};">${type === "Non renseigné" ? "" : type}</td><td style="padding:10px;border:1px solid #e5e7eb;">${client}</td><td style="padding:10px;border:1px solid #e5e7eb;text-align:center;">${pax}</td><td style="padding:10px;border:1px solid #e5e7eb;text-align:center;">${reunion}</td><td style="padding:10px;border:1px solid #e5e7eb;">${calculerDureeDepuisDates_(s, e)}</td><td style="padding:10px;border:1px solid #e5e7eb;">${espaces.join("<br>")}</td><td style="padding:10px;border:1px solid #e5e7eb;">${responsable}</td></tr>`;
  });

  html += `</table></div></div></div>`;
  return { destinataire, subject, htmlBody: html };
}

function construireRecapFiltre(typeFiltre) {
  let rows = [["Event","Date","Prestataire"]];

  eventsGroupes.forEach(g => {
    const desc = g.events[0].description || "";

    let ok = false;

    if (typeFiltre === "ONET") ok = /ONET/i.test(desc);
    if (typeFiltre === "CJ") ok = /CJ/i.test(desc);
    if (typeFiltre === "TECHNIQUE") ok = /TECHNIQUE/i.test(desc);

    if (ok) {
      rows.push([g.nom, g.debutMin, typeFiltre]);
    }
  });

  return { html:"", rows };
}

function exporterPDF(html) {
  const blob = Utilities.newBlob(html, "text/html", "export.html");

  const pdf = blob.getAs("application/pdf");

  const file = DriveApp.createFile(pdf).setName("Export Planning.pdf");

  return file.getUrl();
}
function lancerExport(form) {
  const type = form.typeExport;
  const format = form.formatExport;

  let data;

  switch (type) {
    case "mensuel":
      data = construireRecapMensuel();
      break;
    case "detail":
      data = construireListeDetaillee();
      break;
    case "espace":
      data = construireRecapParEspace();
      break;
    case "responsable":
      data = construireRecapResponsable();
      break;
    case "onet":
      data = construireRecapFiltre("ONET");
      break;
    case "cj":
      data = construireRecapFiltre("CJ");
      break;
    case "tech":
      data = construireRecapFiltre("TECHNIQUE");
      break;
  }

  if (format === "pdf") {
    return exporterPDF(data.html);
  } else {
    return exporterExcel(data.rows);
  }
}
function exporterExcel(rows) {
  const ss = SpreadsheetApp.create("Export Planning");
  const sheet = ss.getActiveSheet();

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // mise en forme rapide
  sheet.getRange(1,1,1,rows[0].length).setFontWeight("bold");

  return ss.getUrl();
}
function construireListeDetaillee() {
  let rows = [
    ["Date début","Date fin","Event","Type","Client","Responsable"]
  ];

  eventsGroupes.forEach(g => {
    rows.push([
      g.debutMin,
      g.finMax,
      g.nom,
      H.detecterType(g.events[0]),
      H.extraireChamp(g.events[0].description, "Client"),
      H.extraireResponsable(g.events[0].description)
    ]);
  });

  return {
    html: "", // optionnel
    rows: rows
  };
}
function createEventFromForm(data) {
  const siteKey = data.site;
  const cfg = getSiteConfig_(siteKey);

  const calendarId = cfg.calendarEvents;
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) {
    throw new Error("Agenda événements introuvable pour " + cfg.label);
  }

  const start = new Date(data.start);
  const end = new Date(data.end);

  const lignes = [
    "Responsable: " + (data.manager || ""),
    "Type: " + (data.type || "B2C"),
    "PAX: " + (data.pax || ""),
    "Espaces: " + ((data.spaces || []).join(", "))
  ];

  if (data.visiteGuidee) {
    lignes.push("Visite guidée: OUI");
    lignes.push("Visite guidée - Date: " + (data.visiteDate || ""));
    lignes.push("Visite guidée - Heure: " + (data.visiteHeure || ""));
    lignes.push("Visite guidée - PAX: " + (data.visitePax || ""));
  } else {
    lignes.push("Visite guidée: NON");
  }

  lignes.push("PMU - Mise en marche borne: " + (data.pmuBorne ? "OUI" : "NON"));
  lignes.push("PMU - Bon à partir: " + (data.pmuBonApartir ? "OUI" : "NON"));
  lignes.push("PMU - Initiation aux paris: " + (data.pmuInitiation ? "OUI" : "NON"));

  if (data.notes) {
    lignes.push("");
    lignes.push(data.notes);
  }

  const description = lignes.join("\n");

  cal.createEvent(
    data.title,
    start,
    end,
    {
      description: description
    }
  );

  synchroniserCalendriersVersFeuille(siteKey);

  return {
    ok: true,
    message: "Événement créé sur " + cfg.label
  };
}

function submitEventForm_(data) {
  if (!data || !data.site || !data.title || !data.start || !data.end) {
    throw new Error("Champs obligatoires manquants.");
  }

  const conflicts = checkConflictsForForm_(data);

  if (conflicts.length) {
    return {
      ok: false,
      conflicts: conflicts
    };
  }

  return createEventFromForm(data);
}

function getSpacesForForm_(siteKey) {
  const sheet = getPlanningSheet_(siteKey);
  const lastCol = sheet.getLastColumn();
  if (lastCol < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const fixedColsCount = getFixedColsCount_(siteKey);
  const idIdx = headers.indexOf("ID_TECH");
  if (idIdx === -1) return [];

  return headers.slice(fixedColsCount, idIdx);
}
function dashboardEnghien() { genererDashboard_("ENGHIEN"); }
function dashboardVincennes() { genererDashboard_("VINCENNES"); }

function conflitsEnghien() { verifierConflits_("ENGHIEN"); }
function conflitsVincennes() { verifierConflits_("VINCENNES"); }

function rapportMensuelCourantEnghien() {
  const now = new Date();
  rapportMensuel_("ENGHIEN", now.getFullYear(), now.getMonth() + 1);
}

function rapportMensuelCourantVincennes() {
  const now = new Date();
  rapportMensuel_("VINCENNES", now.getFullYear(), now.getMonth() + 1);
}

function rapportAnnuelCourantEnghien() {
  rapportAnnuel_("ENGHIEN", new Date().getFullYear());
}

function rapportAnnuelCourantVincennes() {
  rapportAnnuel_("VINCENNES", new Date().getFullYear());
}

// =======================
// OUTILS GÉNÉRAUX
// =======================
function createEventFromForm(data) {
  const siteKey = data.site;
  const cfg = getSiteConfig_(siteKey);

  const calendarId = cfg.calendarEvents;
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) {
    throw new Error("Agenda événements introuvable pour " + cfg.label);
  }

  const start = new Date(data.start);
  const end = new Date(data.end);

  const lignes = [
    "Responsable: " + (data.manager || ""),
    "Type: " + (data.type || "B2C"),
    "PAX: " + (data.pax || ""),
    "Espaces: " + ((data.spaces || []).join(", "))
  ];

  if (data.visiteGuidee) {
    lignes.push("Visite guidée: OUI");
    lignes.push("Visite guidée - Date: " + (data.visiteDate || ""));
    lignes.push("Visite guidée - Heure: " + (data.visiteHeure || ""));
    lignes.push("Visite guidée - PAX: " + (data.visitePax || ""));
  } else {
    lignes.push("Visite guidée: NON");
  }

  lignes.push("PMU - Mise en marche borne: " + (data.pmuBorne ? "OUI" : "NON"));
  lignes.push("PMU - Bon à partir: " + (data.pmuBonApartir ? "OUI" : "NON"));
  lignes.push("PMU - Initiation aux paris: " + (data.pmuInitiation ? "OUI" : "NON"));

  if (data.notes) {
    lignes.push("");
    lignes.push(data.notes);
  }

  const description = lignes.join("\n");

  cal.createEvent(
    data.title,
    start,
    end,
    {
      description: description
    }
  );

  synchroniserCalendriersVersFeuille(siteKey);

  return {
    ok: true,
    message: "Événement créé sur " + cfg.label
  };
}

function submitEventForm_(data) {
  if (!data || !data.site || !data.title || !data.start || !data.end) {
    throw new Error("Champs obligatoires manquants.");
  }

  if (new Date(data.end) <= new Date(data.start)) {
    throw new Error("La date de fin doit être après la date de début.");
  }

  if (data.visiteGuidee) {
    if (!data.visiteDate || !data.visiteHeure) {
      throw new Error("Merci de renseigner la date et l'heure de la visite guidée.");
    }
  }

  const conflicts = checkConflictsForForm_(data);

  if (conflicts.length) {
    return {
      ok: false,
      conflicts: conflicts
    };
  }

  return createEventFromForm(data);
}
function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSiteConfig_(siteKey) {
  const cfg = SITES[siteKey];
  if (!cfg) throw new Error("Site inconnu : " + siteKey);
  return cfg;
}

function ensureSheetByName_(name) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function getPlanningSheet_(siteKey) {
  return ensureSheetByName_(getSiteConfig_(siteKey).sheetName);
}

function getLogSheet_(siteKey) {
  const sh = ensureSheetByName_(getSiteConfig_(siteKey).logSheetName);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "Horodatage",
      "Action",
      "Date",
      "Heure",
      "Événement",
      "Type",
      "ID_TECH",
      "Détail"
    ]);
  }
  return sh;
}

function getDashboardSheet_(siteKey) {
  return ensureSheetByName_(getSiteConfig_(siteKey).dashboardSheetName);
}

function getAlertsSheet_(siteKey) {
  return ensureSheetByName_(getSiteConfig_(siteKey).alertsSheetName);
}

function normalizeText_(str) {
  return String(str || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function cleanText_(str) {
  return String(str || "").replace(/\s+/g, " ").trim();
}

function cleanHtml_(str) {
  return String(str || "").replace(/<\/?[^>]+(>|$)/g, "");
}

function formatDayKey_(date) {
  return Utilities.formatDate(date, TZ, "yyyyMMdd");
}

function formatDateKey_(date) {
  return Utilities.formatDate(date, TZ, "yyyy-MM-dd");
}

function formatDateFr_(date) {
  return Utilities.formatDate(date, TZ, "dd/MM/yyyy");
}

function formatMonthKey_(date) {
  return Utilities.formatDate(date, TZ, "yyyy-MM");
}

function formatWeekKey_(date) {
  return Utilities.formatDate(date, TZ, "YYYY-'S'ww");
}

function columnToLetter_(column) {
  let temp = "";
  let letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function parseDateFlexible_(value) {
  if (value instanceof Date) return new Date(value);

  const s = String(value || "").trim();
  if (!s) return null;

  const mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (mIso) return new Date(Number(mIso[1]), Number(mIso[2]) - 1, Number(mIso[3]));

  const mFr = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (mFr) return new Date(Number(mFr[3]), Number(mFr[2]) - 1, Number(mFr[1]));

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function getSyncRange_() {
  const year = new Date().getFullYear();
  return {
    year: year,
    start: new Date(year, 0, 1),
    end: new Date(year, 11, 31, 23, 59, 59)
  };
}

function getYearRange_(year) {
  return {
    start: new Date(year, 0, 1),
    end: new Date(year, 11, 31, 23, 59, 59)
  };
}

function getMonthRange_(year, month) {
  return {
    start: new Date(year, month - 1, 1),
    end: new Date(year, month, 0, 23, 59, 59)
  };
}

// =======================
// TEST AGENDAS
// =======================

function testerTousLesAgendas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "TEST_AGENDAS";
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
  } else {
    sh.clear();
  }

  const headers = [
    "Site",
    "Type agenda",
    "ID agenda",
    "Test CalendarApp",
    "Test Calendar API",
    "Diagnostic",
    "Détail erreur"
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const results = [];
  const start = new Date();
  const end = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);

  Object.keys(SITES).forEach(function(siteKey) {
    const cfg = SITES[siteKey];

    [
      { type: "Courses", id: cfg.calendarCourses },
      { type: "Événements", id: cfg.calendarEvents }
    ].forEach(function(item) {
      const calendarId = String(item.id || "").trim();

      let calAppStatus = "NON TESTÉ";
      let apiStatus = "NON TESTÉ";
      let diagnostic = "";
      let detail = "";

      if (!calendarId) {
        calAppStatus = "KO";
        apiStatus = "KO";
        diagnostic = "ID agenda vide";
        detail = "Aucun identifiant renseigné dans la configuration.";
      } else {
        try {
          const cal = CalendarApp.getCalendarById(calendarId);
          calAppStatus = cal ? "OK" : "KO";
        } catch (e) {
          calAppStatus = "KO";
          detail += "[CalendarApp] " + e.message + " ";
        }

        try {
          const res = Calendar.Events.list(calendarId, {
            timeMin: start.toISOString(),
            timeMax: end.toISOString(),
            singleEvents: true,
            maxResults: 5
          });
          apiStatus = "OK";
          detail += "[API] Lecture OK, " + ((res.items || []).length) + " événement(s) sur 7 jours. ";
        } catch (e) {
          apiStatus = "KO";
          detail += "[API] " + e.message + " ";
        }

        if (calAppStatus === "OK" && apiStatus === "OK") {
          diagnostic = "OK";
        } else if (calAppStatus === "KO" && apiStatus === "KO") {
          diagnostic = "Agenda introuvable, ID faux ou non partagé";
        } else if (calAppStatus === "OK" && apiStatus === "KO") {
          diagnostic = "CalendarApp OK, API KO : vérifier l’activation de Google Calendar API";
        } else if (calAppStatus === "KO" && apiStatus === "OK") {
          diagnostic = "Cas atypique, vérifier les autorisations du projet";
        }
      }

      results.push([
        cfg.label,
        item.type,
        calendarId,
        calAppStatus,
        apiStatus,
        diagnostic,
        detail.trim()
      ]);
    });
  });

  if (results.length) {
    sh.getRange(2, 1, results.length, headers.length).setValues(results);
  }

  sh.autoResizeColumns(1, headers.length);
  sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sh.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("Test des agendas terminé.\n\nConsulte l’onglet TEST_AGENDAS.");
}

// =======================
// EXTRACTIONS DESCRIPTION
// =======================

function getProjectManager_(event) {
  const descOriginale = cleanHtml_(event.description || "");
  if (!descOriginale.trim()) return "";

  const patterns = [
    /(?:^|\n|\r)\s*(?:responsables?|resp)(?:\s*setf)?\s*[:\-]?\s*(.+)/i,
    /(?:^|\n|\r)\s*(?:gestion|gestionnaire)\s*[:\-]?\s*(.+)/i,
    /(?:^|\n|\r)\s*(?:chef de projet|cp)\s*[:\-]?\s*(.+)/i,
    /(?:^|\n|\r)\s*(?:referent|référent)\s*[:\-]?\s*(.+)/i
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = descOriginale.match(patterns[i]);
    if (match && match[1]) {
      return cleanText_(
        match[1]
          .split(/\n|\r/)[0]
          .replace(/\s*\|\s*.*/, "")
      );
    }
  }

  return "";
}

function getB2BFlag_(event) {
  const desc = String(event.description || "").toLowerCase();
  if (!desc) return "B2C";
  if (desc.includes("b2b")) return "B2B";
  if (desc.includes("b2c")) return "B2C";
  return "B2C";
}

function getPaxFromDescription_(event) {
  const desc = cleanHtml_(event.description || "");
  if (!desc) return "";
  const lower = desc.toLowerCase();

  const patterns = [
    /\b(?:pax|participants?|pers|personnes?|jauge)\s*[:\-]?\s*(\d{1,5})/i,
    /(\d{1,5})\s*(?:pax|participants?|pers|personnes?|jauge)/i
  ];

  for (let i = 0; i < patterns.length; i++) {
    const m = lower.match(patterns[i]);
    if (m && m[1]) {
      const n = parseInt(m[1], 10);
      if (!isNaN(n)) return n;
    }
  }
  return "";
}

// =======================
// CALENDAR API
// =======================

function getAllEventsRaw_(calendarId, start, end) {
  const items = [];
  let pageToken = null;

  do {
    const res = Calendar.Events.list(calendarId, {
      timeMin: start.toISOString(),
      timeMax: end.toISOString(),
      singleEvents: true,
      showDeleted: false,
      maxResults: 2500,
      pageToken: pageToken
    });

    if (res.items && res.items.length) {
      for (let i = 0; i < res.items.length; i++) items.push(res.items[i]);
    }

    pageToken = res.nextPageToken;
  } while (pageToken);

  return items;
}

function getAllEventsForSite_(siteKey, start, end) {
  const cfg = getSiteConfig_(siteKey);
  return {
    courseEvents: getAllEventsRaw_(cfg.calendarCourses, start, end),
    eventEvents: getAllEventsRaw_(cfg.calendarEvents, start, end)
  };
}

// =======================
// ESPACES / RESSOURCES
// =======================

function getDisplayResourceLabel_(raw) {
  const s = String(raw || "").trim();

  const patterns = [
    /SALON HEMINGWAY/i,
    /GRAND HALL/i,
    /HALL DES BALANCES/i,
    /LA ROTONDE/i,
    /BORD DE PISTE/i,
    /LES EXT[ÉE]RIEURS/i,
    /PARKING DES VANS/i,
    /PARKING GRAND PUBLIC/i,
    /PARKING GRAN D PUBLIC/i,
    /PARKING PROPRIETAIRE/i,
    /RESTAURANT PANORAMIQUE/i,
    /LA TERRASSE/i,
    /TRIBUNES?/i,
    /SALLE DES COMMISSAIRES/i,
    /SALLE DES FAIENCES/i,
    /ESPACE LOUNGE DES BALAN/i,
    /LES ÉCURIES/i,
    /LES ECURIES/i
  ];

  for (let i = 0; i < patterns.length; i++) {
    const m = s.match(patterns[i]);
    if (m) {
      return m[0]
        .replace(/GRAN D/i, "GRAND")
        .replace(/EXTERIEURS/i, "EXTÉRIEURS")
        .replace(/ECURIES/i, "ÉCURIES")
        .toUpperCase();
    }
  }

  return s
    .replace(/^\([^)]*\)-/i, "")
    .replace(/Hippodrome (Enghien-Soisy|Paris-Vincennes)-?/i, "")
    .replace(/\(\d+\)/g, "")
    .replace(/\b(RDC|1|2|3)-/g, "")
    .replace(/Intérieur|Extérieur/gi, "")
    .replace(/-+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function extractResourcesFromEvents_(events) {
  const map = new Map();

  for (let i = 0; i < events.length; i++) {
    const attendees = events[i].attendees || [];

    for (let g = 0; g < attendees.length; g++) {
      const attendee = attendees[g];
      if (attendee.resource !== true) continue;

      const raw = String(attendee.displayName || attendee.email || "").trim();
      const norm = normalizeText_(raw);
      if (!norm) continue;

      if (!map.has(norm)) {
        map.set(norm, getDisplayResourceLabel_(raw));
      }
    }
  }

  return map;
}

function getDynamicSpaces_(courseEvents, eventEvents) {
  const allEvents = courseEvents.concat(eventEvents);
  const map = extractResourcesFromEvents_(allEvents);

  return Array.from(map.entries())
    .map(function(entry) {
      return { norm: entry[0], label: entry[1] };
    })
    .sort(function(a, b) {
      return a.label.localeCompare(b.label, "fr", { sensitivity: "base" });
    });
}

function buildSpaceIndex_(siteKey, spaces) {
  const index = {};
  const fixedColsCount = getFixedColsCount_(siteKey);

  for (let i = 0; i < spaces.length; i++) {
    index[spaces[i].norm] = fixedColsCount + i;
  }
  return index;
}

function getEventResourceNorms_(event) {
  const norms = [];
  const seen = {};

  const attendees = event.attendees || [];
  for (let i = 0; i < attendees.length; i++) {
    const attendee = attendees[i];
    if (attendee.resource !== true) continue;

    const norm = normalizeText_(attendee.displayName || attendee.email);
    if (!norm || seen[norm]) continue;

    seen[norm] = true;
    norms.push(norm);
  }

  return norms;
}

// =======================
// COURSES
// =======================

function detectCourseSessionType_(event, startDate) {
  const title = String(event.summary || "").toUpperCase();

  if (title.includes("SEMI-NOCTURNE") || title.includes("SEMI NOCTURNE")) return "Semi-nocturne";
  if (title.includes("NOCTURNE")) return "Nocturne";
  if (title.includes("DIURNE")) return "Diurne";

  const hour = Number(Utilities.formatDate(startDate, TZ, "H"));
  if (hour >= 20) return "Nocturne";
  if (hour >= 17) return "Semi-nocturne";
  return "Diurne";
}

function buildCourseTimeLabel_(event, startDate, endDate, isAllDay) {
  if (isAllDay) return "";
  const typeLabel = detectCourseSessionType_(event, startDate);
  const hDebut = Utilities.formatDate(startDate, TZ, "HH:mm");
  const hFin = Utilities.formatDate(endDate, TZ, "HH:mm");
  return typeLabel + " | " + hDebut + " - " + hFin;
}

function getSiteLabelForReference_(siteKey) {
  return siteKey === "ENGHIEN" ? "Enghien" : "Vincennes";
}

function getGroupe1CoursesForDate_(siteKey, dateObj) {
  const siteLabel = getSiteLabelForReference_(siteKey);
  const dateKey = Utilities.formatDate(dateObj, TZ, "yyyy-MM-dd");

  return GROUPE_1_REFERENCE
    .filter(function(item) {
      return item.hippodrome === siteLabel && item.date2026 === dateKey;
    })
    .map(function(item) {
      return item.course;
    });
}

function getGroupe_1CoursesForDateSafe_(siteKey, dateObj) {
  try {
    return getGroupe1CoursesForDate_(siteKey, dateObj);
  } catch (e) {
    return [];
  }
}

function getCoursePhareLabel_(siteKey, dateObj) {
  const courses = getGroupe_1CoursesForDateSafe_(siteKey, dateObj);
  return courses.length ? courses.join(" | ") : "";
}

function getCourseLabelForPlanning_(siteKey, dateObj) {
  const coursePhare = getCoursePhareLabel_(siteKey, dateObj);
  return coursePhare || "Réunion de courses";
}

function isCoursePhare_(label) {
  const txt = String(label || "").trim();
  if (!txt) return false;

  return GROUPE_1_REFERENCE.some(function(item) {
    return item.course === txt;
  });
}

function isCourseRow_(siteKey, eventText, typeValue) {
  return String(typeValue || "").trim() === "Course";
}

// =======================
// CONSTRUCTION DU PLANNING
// =======================

function parseEventDates_(event) {
  const isAllDay = !!(event.start && event.start.date && !event.start.dateTime);

  let startDate = null;
  let endDate = null;

  if (isAllDay) {
    startDate = new Date(event.start.date);
    endDate = new Date(event.end.date);
    endDate = new Date(endDate.getTime() - 86400000);
  } else {
    startDate = new Date(event.start.dateTime);
    endDate = new Date(event.end.dateTime);
  }

  return { startDate: startDate, endDate: endDate, isAllDay: isAllDay };
}

function createBaseRowsByDay_(year) {
  const rows = [];
  for (let d = new Date(year, 0, 1); d <= new Date(year, 11, 31); d.setDate(d.getDate() + 1)) {
    const dateObj = new Date(d);
    rows.push({
      dayKey: formatDayKey_(dateObj),
      dateObj: dateObj,
      items: []
    });
  }
  return rows;
}

function getFixedHeaders_(siteKey) {
  return [
    "Date N",
    "Evénement + Heure",
    "Gestion",
    "Type",
    "PAX"
  ];
}

function getFixedColsCount_(siteKey) {
  return getFixedHeaders_(siteKey).length;
}

function getGestionColumnIndex_(siteKey) {
  return 2;
}

function getTypeColumnIndex_(siteKey) {
  return 3;
}

function getHeaders_(siteKey, spaces) {
  return getFixedHeaders_(siteKey).concat(
    spaces.map(function(s) { return s.label; }),
    [TRAILING_HEADER]
  );
}

function buildEventRow_(siteKey, event, type, dateObj, spacesCount, spaceIndex) {
  const parsed = parseEventDates_(event);
  const startDate = parsed.startDate;
  const endDate = parsed.endDate;
  const isAllDay = parsed.isAllDay;

  const fixedColsCount = getFixedColsCount_(siteKey);
  const row = new Array(fixedColsCount + spacesCount + 1).fill("");

  const resourceNorms = getEventResourceNorms_(event);
  const manager = getProjectManager_(event);
  const b2b = getB2BFlag_(event);
  const pax = getPaxFromDescription_(event);

  const label = (type === "course")
    ? getCourseLabelForPlanning_(siteKey, dateObj)
    : String(event.summary || "");

  const heure = (type === "course")
    ? buildCourseTimeLabel_(event, startDate, endDate, isAllDay)
    : "";

  row[0] = new Date(dateObj);
  row[1] = heure ? (label + " | " + heure) : label;
  row[2] = manager;
  row[3] = (type === "course") ? "Course" : b2b;
  row[4] = pax;

  for (let i = 0; i < resourceNorms.length; i++) {
    const col = spaceIndex[resourceNorms[i]];
    if (col !== undefined) row[col] = "X";
  }

  row[row.length - 1] =
    (event.id || event.iCalUID || "NO_ID") + "_" +
    Utilities.formatDate(dateObj, TZ, "yyyyMMdd");

  return row;
}

function appendEventsToDays_(siteKey, events, type, dayIndex, spacesCount, spaceIndex) {
  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    const parsed = parseEventDates_(event);
    const startDate = parsed.startDate;
    const endDate = parsed.endDate;

    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
      const key = formatDayKey_(d);
      const day = dayIndex[key];
      if (!day) continue;
      day.items.push(buildEventRow_(siteKey, event, type, new Date(d), spacesCount, spaceIndex));
    }
  }
}

function buildCalendarRowsAndEntries_(siteKey, courseEvents, eventEvents, spaces, year) {
  const baseDays = createBaseRowsByDay_(year);
  const dayIndex = {};
  const spaceIndex = buildSpaceIndex_(siteKey, spaces);
  const fixedColsCount = getFixedColsCount_(siteKey);

  for (let i = 0; i < baseDays.length; i++) {
    dayIndex[baseDays[i].dayKey] = baseDays[i];
  }

  appendEventsToDays_(siteKey, courseEvents, "course", dayIndex, spaces.length, spaceIndex);
  appendEventsToDays_(siteKey, eventEvents, "event", dayIndex, spaces.length, spaceIndex);

  const rows = [];
  const entries = {};

  for (let i = 0; i < baseDays.length; i++) {
    const day = baseDays[i];

    if (!day.items.length) {
      const emptyRow = new Array(fixedColsCount + spaces.length + 1).fill("");
      emptyRow[0] = day.dateObj;
      emptyRow[emptyRow.length - 1] = "DAY_" + day.dayKey;
      rows.push(emptyRow);

      entries[emptyRow[emptyRow.length - 1]] = {
        date: formatDateKey_(day.dateObj),
        evenement: "",
        heure: "",
        type: "",
        gestion: "",
        b2b: "",
        pax: "",
        spaces: []
      };
      continue;
    }

    day.items.sort(function(a, b) {
      const dateDiff = new Date(a[0]).getTime() - new Date(b[0]).getTime();
      if (dateDiff !== 0) return dateDiff;

      const labelA = String(a[1] || "");
      const labelB = String(b[1] || "");
      return labelA.localeCompare(labelB);
    });

    for (let j = 0; j < day.items.length; j++) {
      const row = day.items[j];
      rows.push(row);

      const id = row[row.length - 1];
      const spacesForEntry = [];

      for (let c = fixedColsCount; c < row.length - 1; c++) {
        if (row[c] === "X") {
          spacesForEntry.push(normalizeText_(spaces[c - fixedColsCount].label));
        }
      }

      entries[id] = {
        date: formatDateKey_(row[0]),
        evenement: String(row[1] || ""),
        heure: "",
        type: String(row[3] || ""),
        gestion: String(row[2] || ""),
        b2b: String(row[3] === "Course" ? "" : row[3] || ""),
        pax: String(row[4] || ""),
        spaces: spacesForEntry
      };
    }
  }

  rows.sort(function(a, b) {
    const dateDiff = new Date(a[0]).getTime() - new Date(b[0]).getTime();
    if (dateDiff !== 0) return dateDiff;
    return String(a[1] || "").localeCompare(String(b[1] || ""));
  });

  return { rows: rows, entries: entries };
}

// =======================
// LECTURE / LOG DELTA
// =======================

function rebuildHeaders_(siteKey, sheet, spaces) {
  sheet.clear();
  const headers = getHeaders_(siteKey, spaces);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.hideColumn(sheet.getRange(1, headers.length));
}

function readExistingEntries_(siteKey, sheet) {
  const result = {};

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return result;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const fixedColsCount = getFixedColsCount_(siteKey);
  const dateIdx = headers.indexOf("Date N");
  const eventIdx = headers.indexOf("Evénement + Heure");
  const gestionIdx = headers.indexOf("Gestion");
  const typeIdx = headers.indexOf("Type");
  const paxIdx = headers.indexOf("PAX");
  const idIdx = headers.indexOf("ID_TECH");

  if (idIdx === -1) return result;

  const dynamicHeaders = headers.slice(fixedColsCount, idIdx);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const id = row[idIdx];
    if (!id) continue;

    const spaces = [];
    for (let j = 0; j < dynamicHeaders.length; j++) {
      if (row[fixedColsCount + j] === "X") {
        spaces.push(normalizeText_(dynamicHeaders[j]));
      }
    }

    result[id] = {
      date: row[dateIdx] instanceof Date ? formatDateKey_(row[dateIdx]) : String(row[dateIdx] || ""),
      evenement: String(row[eventIdx] || ""),
      heure: "",
      type: String((typeIdx > -1 ? row[typeIdx] : "") || ""),
      gestion: String((gestionIdx > -1 ? row[gestionIdx] : "") || ""),
      b2b: String((typeIdx > -1 && row[typeIdx] !== "Course" ? row[typeIdx] : "") || ""),
      pax: String((paxIdx > -1 ? row[paxIdx] : "") || ""),
      spaces: spaces
    };
  }

  return result;
}

function stringifyEntry_(entry) {
  return JSON.stringify({
    date: entry.date || "",
    evenement: entry.evenement || "",
    heure: entry.heure || "",
    type: entry.type || "",
    gestion: entry.gestion || "",
    b2b: entry.b2b || "",
    pax: entry.pax || "",
    spaces: (entry.spaces || []).slice().sort()
  });
}

function buildLogEntry_(action, rowOrEntry, detail) {
  let dateValue = "";
  let heureValue = "";
  let eventValue = "";
  let typeValue = "";
  let idValue = "";

  if (Array.isArray(rowOrEntry)) {
    dateValue = rowOrEntry[0] || "";
    eventValue = rowOrEntry[1] || "";
    heureValue = "";
    typeValue = rowOrEntry[3] || "";
    idValue = rowOrEntry[rowOrEntry.length - 1] || "";
  } else {
    dateValue = rowOrEntry.date || "";
    eventValue = rowOrEntry.evenement || "";
    heureValue = rowOrEntry.heure || "";
    typeValue = rowOrEntry.type || "";
    idValue = rowOrEntry.id || "";
  }

  return [
    new Date(),
    action,
    dateValue,
    heureValue,
    eventValue,
    typeValue,
    idValue,
    detail || ""
  ];
}

function buildUpdateDetail_(oldEntry, newEntry) {
  const changes = [];

  if ((oldEntry.evenement || "") !== (newEntry.evenement || "")) {
    changes.push("événement: '" + (oldEntry.evenement || "") + "' → '" + (newEntry.evenement || "") + "'");
  }
  if ((oldEntry.type || "") !== (newEntry.type || "")) {
    changes.push("type: '" + (oldEntry.type || "") + "' → '" + (newEntry.type || "") + "'");
  }
  if ((oldEntry.gestion || "") !== (newEntry.gestion || "")) {
    changes.push("gestion: '" + (oldEntry.gestion || "") + "' → '" + (newEntry.gestion || "") + "'");
  }
  if ((oldEntry.b2b || "") !== (newEntry.b2b || "")) {
    changes.push("B2B/B2C: '" + (oldEntry.b2b || "") + "' → '" + (newEntry.b2b || "") + "'");
  }
  if ((oldEntry.pax || "") !== (newEntry.pax || "")) {
    changes.push("PAX: '" + (oldEntry.pax || "") + "' → '" + (newEntry.pax || "") + "'");
  }

  const oldSpaces = (oldEntry.spaces || []).slice().sort().join(", ");
  const newSpaces = (newEntry.spaces || []).slice().sort().join(", ");
  if (oldSpaces !== newSpaces) {
    changes.push("espaces: '" + oldSpaces + "' → '" + newSpaces + "'");
  }

  return changes.join(" | ");
}

function buildDeltaLogEntries_(oldEntries, newEntries) {
  const logs = [];
  const seen = {};

  for (const id in oldEntries) {
    seen[id] = true;

    if (!newEntries[id]) {
      logs.push(buildLogEntry_("SUPPRESSION", {
        date: oldEntries[id].date,
        evenement: oldEntries[id].evenement,
        heure: oldEntries[id].heure,
        type: oldEntries[id].type,
        id: id
      }, ""));
      continue;
    }

    if (stringifyEntry_(oldEntries[id]) !== stringifyEntry_(newEntries[id])) {
      logs.push(buildLogEntry_("MODIFICATION", {
        date: newEntries[id].date,
        evenement: newEntries[id].evenement,
        heure: newEntries[id].heure,
        type: newEntries[id].type,
        id: id
      }, buildUpdateDetail_(oldEntries[id], newEntries[id])));
    }
  }

  for (const id in newEntries) {
    if (seen[id]) continue;

    logs.push(buildLogEntry_("AJOUT", {
      date: newEntries[id].date,
      evenement: newEntries[id].evenement,
      heure: newEntries[id].heure,
      type: newEntries[id].type,
      id: id
    }, ""));
  }

  return logs;
}

function writeSyncLog_(siteKey, entries) {
  if (!entries || !entries.length) return;
  const logSheet = getLogSheet_(siteKey);
  logSheet.getRange(logSheet.getLastRow() + 1, 1, entries.length, entries[0].length)
    .setValues(entries);
}

// =======================
// SYNCHRO PRINCIPALE
// =======================

function synchroniserCalendriersVersFeuille(siteKey) {
  const sheet = getPlanningSheet_(siteKey);
  const oldEntries = readExistingEntries_(siteKey, sheet);

  const syncRange = getSyncRange_();
  const allEvents = getAllEventsForSite_(siteKey, syncRange.start, syncRange.end);

  const spaces = getDynamicSpaces_(allEvents.courseEvents, allEvents.eventEvents);
  rebuildHeaders_(siteKey, sheet, spaces);

  const built = buildCalendarRowsAndEntries_(
    siteKey,
    allEvents.courseEvents,
    allEvents.eventEvents,
    spaces,
    syncRange.year
  );

  const headers = getHeaders_(siteKey, spaces);
  const expectedCols = headers.length;
  const actualCols = built.rows.length ? built.rows[0].length : 0;

  Logger.log("=== SYNCHRO " + siteKey + " ===");
  Logger.log("Headers count = " + expectedCols);
  Logger.log("Rows count = " + built.rows.length);
  Logger.log("Row length = " + actualCols);

  if (built.rows.length) {
    if (expectedCols !== actualCols) {
      Logger.log("HEADERS DETAIL = " + JSON.stringify(headers));
      Logger.log("ROW DETAIL = " + JSON.stringify(built.rows[0]));
      throw new Error(
        "Mismatch colonnes pour " + siteKey +
        " : headers=" + expectedCols +
        " / rows=" + actualCols
      );
    }

    sheet.getRange(2, 1, built.rows.length, actualCols).setValues(built.rows);
  }

  appliquerMiseEnPage(siteKey);
  nettoyerAffichageDates_(siteKey);
  masquerLignesPassees_(siteKey);
  styliserCoursesPhares_(siteKey);

  const logEntries = buildDeltaLogEntries_(oldEntries, built.entries);
  writeSyncLog_(siteKey, logEntries);

  return {
    site: getSiteConfig_(siteKey).label,
    rows: built.rows.length,
    spaces: spaces.length,
    added: logEntries.filter(function(r) { return r[1] === "AJOUT"; }).length,
    updated: logEntries.filter(function(r) { return r[1] === "MODIFICATION"; }).length,
    deleted: logEntries.filter(function(r) { return r[1] === "SUPPRESSION"; }).length
  };
}

// =======================
// MASQUAGE / AFFICHAGE
// =======================

function masquerLignesPassees_(siteKey) {
  const sheet = getPlanningSheet_(siteKey);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  sheet.showRows(2, lastRow - 1);

  const idTechCol = lastCol;
  const ids = sheet.getRange(2, idTechCol, lastRow - 1, 1).getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const rowsToHide = [];

  for (let i = 0; i < ids.length; i++) {
    const row = i + 2;
    const id = String(ids[i][0] || "").trim();
    if (!id) continue;

    const match = id.match(/(\d{8})$/);
    if (!match) continue;

    const ymd = match[1];
    const rowDate = new Date(
      Number(ymd.substring(0, 4)),
      Number(ymd.substring(4, 6)) - 1,
      Number(ymd.substring(6, 8))
    );
    rowDate.setHours(0, 0, 0, 0);

    if (rowDate < today) {
      rowsToHide.push(row);
    }
  }

  if (!rowsToHide.length) return;

  let start = rowsToHide[0];
  let prev = rowsToHide[0];

  for (let i = 1; i < rowsToHide.length; i++) {
    const current = rowsToHide[i];

    if (current === prev + 1) {
      prev = current;
    } else {
      sheet.hideRows(start, prev - start + 1);
      start = current;
      prev = current;
    }
  }

  sheet.hideRows(start, prev - start + 1);
}

function masquerLignesPasseesTousSites_() {
  masquerLignesPassees_("ENGHIEN");
  masquerLignesPassees_("VINCENNES");
}

function afficherToutesLesLignes_(siteKey) {
  const sheet = getPlanningSheet_(siteKey);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.showRows(2, lastRow - 1);
  }
}

function afficherToutesLesLignesTousSites_() {
  afficherToutesLesLignes_("ENGHIEN");
  afficherToutesLesLignes_("VINCENNES");
}

function nettoyerAffichageDates_(siteKey) {
  const sheet = getPlanningSheet_(siteKey);
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 1);
  const values = range.getValues();

  let previousDateKey = null;

  for (let i = 0; i < values.length; i++) {
    const cell = values[i][0];
    if (!(cell instanceof Date)) continue;

    const currentKey = Utilities.formatDate(cell, TZ, "yyyyMMdd");

    if (currentKey === previousDateKey) {
      values[i][0] = "";
    } else {
      previousDateKey = currentKey;
    }
  }

  range.setValues(values);
}

// =======================
// STYLISATION COURSES PHARES
// =======================

function styliserCoursesPhares_(siteKey) {
  const sheet = getPlanningSheet_(siteKey);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = i + 2;
    const text = String(values[i][0] || "");
    if (!text) continue;

    const parts = text.split("|");
    const courseName = parts[0].trim();

    if (!isCoursePhare_(courseName)) {
      sheet.getRange(row, 2)
        .setFontFamily("Arial")
        .setFontStyle("normal")
        .setFontSize(11);
      continue;
    }

    const builder = SpreadsheetApp.newRichTextValue().setText(text);

    builder.setTextStyle(
      0,
      courseName.length,
      SpreadsheetApp.newTextStyle()
        .setFontFamily("Oswald")
        .setItalic(true)
        .setFontSize(10)
        .build()
    );

    builder.setTextStyle(
      courseName.length,
      text.length,
      SpreadsheetApp.newTextStyle()
        .setFontFamily("Arial")
        .setItalic(false)
        .setFontSize(11)
        .build()
    );

    sheet.getRange(row, 2).setRichTextValue(builder.build());
  }
}

// =======================
// MISE EN PAGE
// =======================

function appliquerMiseEnPage(siteKey) {
  const sheet = getPlanningSheet_(siteKey);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0) return;

  const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
  const fixedColsCount = getFixedColsCount_(siteKey);

  sheet.clearFormats();
  sheet.clearNotes();
  try { dataRange.breakApart(); } catch (e) {}

  dataRange
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(11)
    .setFontWeight("normal");

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 520);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 80);

  const espaceStartCol = fixedColsCount + 1;
  const espaceCount = Math.max(0, lastCol - fixedColsCount - 1);

  for (let i = espaceStartCol; i < lastCol; i++) {
    sheet.setColumnWidth(i, 35);
  }

  sheet.setRowHeight(1, 36);
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground("#ffffff")
    .setFontWeight("bold")
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setTextRotation(0);

  if (espaceCount > 0) {
    sheet.getRange(1, espaceStartCol, 1, espaceCount)
      .setTextRotation(90)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(false);

    sheet.getRange(2, espaceStartCol, Math.max(1, lastRow - 1), espaceCount)
      .setWrap(false)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");
  }

  if (lastRow > 1) {
    sheet.getRange(1, 3, 1, 1).setBackground("#eeeeee");
    sheet.getRange(1, 4, 1, 1).setBackground("#e6f3ff");
    sheet.getRange(1, 5, 1, 1).setBackground("#f7f7d9");

    const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const rangeListEvents = [];
    const rangeListCourses = [];
    const rangeListEmpty = [];

    const dateCounts = {};
    for (let i = 0; i < dates.length; i++) {
      if (dates[i][0] instanceof Date) {
        const key = formatDayKey_(dates[i][0]);
        dateCounts[key] = (dateCounts[key] || 0) + 1;
      }
    }

    const boldRows = [];
    const solidTopRows = [];
    const dottedTopRows = [];
    const lastColLetter = columnToLetter_(lastCol);

    let previousDayKey = null;

    for (let i = 0; i < dates.length; i++) {
      const row = i + 2;

      let currentDayKey = null;
      if (dates[i][0] instanceof Date) {
        currentDayKey = formatDayKey_(dates[i][0]);
      }

      const currentType = String(sheet.getRange(row, 4).getValue() || "").trim();
      const blockEndCol = 5;

      if (currentType === "Course") {
        rangeListCourses.push("A" + row + ":" + columnToLetter_(blockEndCol) + row);
      } else if (currentType === "B2B" || currentType === "B2C") {
        rangeListEvents.push("A" + row + ":" + columnToLetter_(blockEndCol) + row);
      } else {
        rangeListEmpty.push("A" + row + ":" + columnToLetter_(blockEndCol) + row);
      }

      if (currentDayKey && (dateCounts[currentDayKey] || 0) > 1) {
        boldRows.push("A" + row + ":" + lastColLetter + row);
      }

      if (i === 0 || currentDayKey !== previousDayKey) {
        solidTopRows.push("A" + row + ":" + lastColLetter + row);
      } else {
        dottedTopRows.push("A" + row + ":" + lastColLetter + row);
      }

      previousDayKey = currentDayKey;
    }

    if (rangeListEvents.length) {
      sheet.getRangeList(rangeListEvents).setBackground("#dff3ff");
    }
    if (rangeListCourses.length) {
      sheet.getRangeList(rangeListCourses).setBackground("#fff1cc");
    }
    if (rangeListEmpty.length) {
      sheet.getRangeList(rangeListEmpty).setBackground("white");
    }

    if (espaceCount > 0) {
      const espacesRange = sheet.getRange(2, espaceStartCol, lastRow - 1, espaceCount);
      espacesRange
        .setBackground("#f4f4f4")
        .setWrap(false)
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center");
    }

    sheet.getRange(2, 1, lastRow - 1, lastCol).setFontWeight("normal");
    if (boldRows.length) {
      sheet.getRangeList(boldRows).setFontWeight("bold");
    }

    sheet.getRange(2, 1, lastRow - 1, lastCol)
      .setBorder(false, null, false, false, false, false);

    if (solidTopRows.length) {
      sheet.getRangeList(solidTopRows).setBorder(
        true, null, null, null, null, null,
        "black",
        SpreadsheetApp.BorderStyle.SOLID
      );
    }

    if (dottedTopRows.length) {
      sheet.getRangeList(dottedTopRows).setBorder(
        true, null, null, null, null, null,
        "#999999",
        SpreadsheetApp.BorderStyle.DOTTED
      );
    }

    sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('[$-fr]dddd d mmmm yyyy');
    sheet.getRange(2, 2, lastRow - 1, fixedColsCount - 1).setWrap(true);
  }

  if (espaceCount > 0) {
    for (let col = espaceStartCol; col < lastCol; col++) {
      sheet.getRange(1, col, lastRow, 1).setBorder(
        null,
        true,
        null,
        true,
        null,
        null,
        "black",
        SpreadsheetApp.BorderStyle.DOUBLE
      );
    }
  }

  dataRange.setBorder(
    true, true, true, true, null, null,
    "black",
    SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.hideColumn(sheet.getRange(1, lastCol));
}

// =======================
// LECTURE DU PLANNING
// =======================

function readPlanningData_(siteKey) {
  const sheet = getPlanningSheet_(siteKey);
  if (!sheet) return { headers: [], rows: [] };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1) return { headers: [], rows: [] };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  const rows = values.map(function(row) {
    const obj = {};
    for (let i = 0; i < headers.length; i++) obj[headers[i]] = row[i];
    return obj;
  });

  return { headers: headers, rows: rows };
}

function getDynamicSpaceHeadersFromSheet_(siteKey, headers) {
  const fixedColsCount = getFixedColsCount_(siteKey);
  const idIdx = headers.indexOf("ID_TECH");
  if (idIdx === -1) return [];
  return headers.slice(fixedColsCount, idIdx);
}

function filterRowsByDate_(rows, startDate, endDate) {
  return rows.filter(function(r) {
    const d = r["Date N"];
    if (!(d instanceof Date)) return false;
    return d >= startDate && d <= endDate;
  });
}

// =======================
// DASHBOARD
// =======================

function genererDashboard_(siteKey) {
  const sh = getDashboardSheet_(siteKey);
  sh.clear();

  const data = readPlanningData_(siteKey);
  const rows = data.rows;
  const headers = data.headers;
  const spaces = getDynamicSpaceHeadersFromSheet_(siteKey, headers);

  const validRows = rows.filter(function(r) {
    return r["Date N"] instanceof Date;
  });

  const courseRows = validRows.filter(function(r) {
    return isCourseRow_(siteKey, r["Evénement + Heure"], r["Type"]);
  });

  const eventRows = validRows.filter(function(r) {
    const hasLabel = String(r["Evénement + Heure"] || "").trim() !== "";
    return hasLabel && !isCourseRow_(siteKey, r["Evénement + Heure"], r["Type"]);
  });

  const occupiedDays = new Set(
    validRows
      .filter(function(r) {
        return cleanText_(r["Evénement + Heure"]);
      })
      .map(function(r) {
        return formatDayKey_(r["Date N"]);
      })
  );

  const byMonth = {};
  const byWeek = {};
  const byManager = {};
  const bySpace = {};

  spaces.forEach(function(s) {
    bySpace[s] = 0;
  });

  validRows.forEach(function(r) {
    const d = r["Date N"];
    byMonth[formatMonthKey_(d)] = (byMonth[formatMonthKey_(d)] || 0) + 1;
    byWeek[formatWeekKey_(d)] = (byWeek[formatWeekKey_(d)] || 0) + 1;

    const manager = cleanText_(r["Gestion"]) || "(vide)";
    byManager[manager] = (byManager[manager] || 0) + 1;

    spaces.forEach(function(s) {
      if (r[s] === "X") bySpace[s] = (bySpace[s] || 0) + 1;
    });
  });

  let row = 1;
  sh.getRange(row++, 1)
    .setValue("DASHBOARD " + getSiteConfig_(siteKey).label.toUpperCase())
    .setFontWeight("bold")
    .setFontSize(14);

  sh.getRange(row, 1, 6, 2).setValues([
    ["Total lignes", validRows.length],
    ["Courses", courseRows.length],
    ["Evénements", eventRows.length],
    ["Jours occupés", occupiedDays.size],
    ["Espaces détectés", spaces.length],
    ["Responsables distincts", Object.keys(byManager).length]
  ]);
  row += 8;

  sh.getRange(row++, 1).setValue("Répartition mensuelle").setFontWeight("bold");
  const monthData = Object.keys(byMonth).sort().map(function(k) {
    return [k, byMonth[k]];
  });
  if (monthData.length) {
    sh.getRange(row, 1, monthData.length, 2).setValues(monthData);
  }
  row += Math.max(2, monthData.length + 2);

  sh.getRange(row++, 1).setValue("Top responsables").setFontWeight("bold");
  const managerData = Object.keys(byManager)
    .map(function(k) { return [k, byManager[k]]; })
    .sort(function(a, b) { return b[1] - a[1]; });

  if (managerData.length) {
    sh.getRange(row, 1, managerData.length, 2).setValues(managerData);
  }
  row += Math.max(2, managerData.length + 2);

  sh.getRange(row++, 1).setValue("Top espaces").setFontWeight("bold");
  const spaceData = Object.keys(bySpace)
    .map(function(k) { return [k, bySpace[k]]; })
    .sort(function(a, b) { return b[1] - a[1]; });

  if (spaceData.length) {
    sh.getRange(row, 1, spaceData.length, 2).setValues(spaceData);
  }

  sh.autoResizeColumns(1, 4);
}

// =======================
// RAPPORTS
// =======================

function buildPeriodSummary_(siteKey, rows) {
  const occupiedDays = new Set(
    rows
      .filter(function(r) { return cleanText_(r["Evénement + Heure"]); })
      .map(function(r) { return formatDayKey_(r["Date N"]); })
  );

  const courses = rows.filter(function(r) {
    return isCourseRow_(siteKey, r["Evénement + Heure"], r["Type"]);
  }).length;

  const events = rows.filter(function(r) {
    const hasLabel = String(r["Evénement + Heure"] || "").trim() !== "";
    return hasLabel && !isCourseRow_(siteKey, r["Evénement + Heure"], r["Type"]);
  }).length;

  return [
    ["Total lignes", rows.length],
    ["Courses", courses],
    ["Evénements", events],
    ["Jours occupés", occupiedDays.size]
  ];
}

function rapportParEspace_(siteKey, espace, dateDebut, dateFin) {
  const sh = ensureSheetByName_("RAPPORT_" + siteKey + "_ESPACES");
  sh.clear();

  const data = readPlanningData_(siteKey);
  const startDate = parseDateFlexible_(dateDebut);
  const endDate = parseDateFlexible_(dateFin);
  if (!startDate || !endDate) throw new Error("Dates invalides");

  const filtered = filterRowsByDate_(data.rows, startDate, endDate).filter(function(r) {
    return r[espace] === "X";
  });

  sh.getRange(1, 1).setValue("RAPPORT ESPACE");
  sh.getRange(2, 1, 4, 2).setValues([
    ["Site", siteKey],
    ["Espace", espace],
    ["Date début", formatDateFr_(startDate)],
    ["Date fin", formatDateFr_(endDate)]
  ]);

  if (filtered.length) {
    const exportRows = filtered.map(function(r) {
      return [
        r["Date N"],
        r["Evénement + Heure"],
        r["Gestion"],
        r["Type"],
        r["PAX"]
      ];
    });
    sh.getRange(7, 1, 1, 5).setValues([["Date", "Evénement + Heure", "Gestion", "Type", "PAX"]]);
    sh.getRange(8, 1, exportRows.length, 5).setValues(exportRows);
    sh.getRange(8, 1, exportRows.length, 1).setNumberFormat('[$-fr]dddd d mmmm yyyy');
  }

  sh.autoResizeColumns(1, 5);
}

function rapportParResponsable_(siteKey, responsable, dateDebut, dateFin) {
  const sh = ensureSheetByName_("RAPPORT_" + siteKey + "_RESPONSABLES");
  sh.clear();

  const data = readPlanningData_(siteKey);
  const startDate = parseDateFlexible_(dateDebut);
  const endDate = parseDateFlexible_(dateFin);
  if (!startDate || !endDate) throw new Error("Dates invalides");

  const respNorm = normalizeText_(responsable);

  const filtered = filterRowsByDate_(data.rows, startDate, endDate)
    .filter(function(r) { return normalizeText_(r["Gestion"]) === respNorm; });

  sh.getRange(1, 1).setValue("RAPPORT RESPONSABLE");
  sh.getRange(2, 1, 4, 2).setValues([
    ["Site", siteKey],
    ["Responsable", responsable],
    ["Date début", formatDateFr_(startDate)],
    ["Date fin", formatDateFr_(endDate)]
  ]);

  if (filtered.length) {
    const dynamicSpaces = getDynamicSpaceHeadersFromSheet_(siteKey, data.headers);
    const exportRows = filtered.map(function(r) {
      const usedSpaces = dynamicSpaces.filter(function(s) { return r[s] === "X"; }).join(", ");
      return [
        r["Date N"],
        r["Evénement + Heure"],
        r["Type"],
        r["PAX"],
        usedSpaces
      ];
    });

    sh.getRange(7, 1, 1, 5).setValues([["Date", "Evénement + Heure", "Type", "PAX", "Espaces"]]);
    sh.getRange(8, 1, exportRows.length, 5).setValues(exportRows);
    sh.getRange(8, 1, exportRows.length, 1).setNumberFormat('[$-fr]dddd d mmmm yyyy');
  }

  sh.autoResizeColumns(1, 5);
}

function rapportMensuel_(siteKey, annee, mois) {
  const sh = ensureSheetByName_("RAPPORT_" + siteKey + "_PERIODES");
  sh.clear();

  const range = getMonthRange_(annee, mois);
  const rows = filterRowsByDate_(readPlanningData_(siteKey).rows, range.start, range.end);

  sh.getRange(1, 1).setValue("RAPPORT MENSUEL");
  sh.getRange(2, 1, 4, 2).setValues([
    ["Site", siteKey],
    ["Année", annee],
    ["Mois", mois],
    ["Période", formatDateFr_(range.start) + " au " + formatDateFr_(range.end)]
  ]);

  const summary = buildPeriodSummary_(siteKey, rows);
  sh.getRange(7, 1, summary.length, 2).setValues(summary);

  if (rows.length) {
    const exportRows = rows.map(function(r) {
      return [
        r["Date N"],
        r["Evénement + Heure"],
        r["Gestion"],
        r["Type"],
        r["PAX"]
      ];
    });
    sh.getRange(14, 1, 1, 5).setValues([["Date", "Evénement + Heure", "Gestion", "Type", "PAX"]]);
    sh.getRange(15, 1, exportRows.length, 5).setValues(exportRows);
    sh.getRange(15, 1, exportRows.length, 1).setNumberFormat('[$-fr]dddd d mmmm yyyy');
  }

  sh.autoResizeColumns(1, 5);
}

function rapportAnnuel_(siteKey, annee) {
  const sh = ensureSheetByName_("RAPPORT_" + siteKey + "_PERIODES");
  sh.clear();

  const range = getYearRange_(annee);
  const rows = filterRowsByDate_(readPlanningData_(siteKey).rows, range.start, range.end);

  sh.getRange(1, 1).setValue("RAPPORT ANNUEL");
  sh.getRange(2, 1, 3, 2).setValues([
    ["Site", siteKey],
    ["Année", annee],
    ["Période", formatDateFr_(range.start) + " au " + formatDateFr_(range.end)]
  ]);

  const summary = buildPeriodSummary_(siteKey, rows);
  sh.getRange(6, 1, summary.length, 2).setValues(summary);

  const byMonth = {};
  rows.forEach(function(r) {
    const d = r["Date N"];
    byMonth[formatMonthKey_(d)] = (byMonth[formatMonthKey_(d)] || 0) + 1;
  });

  const monthRows = Object.keys(byMonth).sort().map(function(k) { return [k, byMonth[k]]; });
  if (monthRows.length) {
    sh.getRange(13, 1, 1, 2).setValues([["Mois", "Lignes"]]);
    sh.getRange(14, 1, monthRows.length, 2).setValues(monthRows);
  }

  sh.autoResizeColumns(1, 5);
}

// =======================
// ALERTES / CONFLITS
// =======================

function verifierConflits_(siteKey) {
  const sh = getAlertsSheet_(siteKey);
  sh.clear();

  const data = readPlanningData_(siteKey);
  const headers = data.headers;
  const rows = data.rows;
  const spaces = getDynamicSpaceHeadersFromSheet_(siteKey, headers);

  const alerts = [];
  const occupancy = {};

  rows.forEach(function(r) {
    const d = r["Date N"];
    if (!(d instanceof Date)) return;

    const keyDay = formatDayKey_(d);
    const id = r["ID_TECH"] || "";
    const type = isCourseRow_(siteKey, r["Evénement + Heure"], r["Type"])
      ? "Course"
      : "Evénement";

    const eventLabel = r["Evénement + Heure"] || "(vide)";

    if (!cleanText_(r["Gestion"]) && cleanText_(r["Evénement + Heure"])) {
      alerts.push([formatDateFr_(d), "GESTION MANQUANTE", eventLabel, type, id]);
    }

    let hasSpace = false;

    spaces.forEach(function(s) {
      if (r[s] === "X") {
        hasSpace = true;
        const occKey = keyDay + "||" + s;
        if (!occupancy[occKey]) occupancy[occKey] = [];
        occupancy[occKey].push(eventLabel + " [" + id + "]");
      }
    });

    if (!hasSpace && cleanText_(r["Evénement + Heure"])) {
      alerts.push([formatDateFr_(d), "ESPACE MANQUANT", eventLabel, type, id]);
    }
  });

  Object.keys(occupancy).forEach(function(k) {
    if (occupancy[k].length > 1) {
      const parts = k.split("||");
      alerts.push([parts[0], "DOUBLE OCCUPATION", parts[1], occupancy[k].join(" | "), ""]);
    }
  });

  sh.getRange(1, 1).setValue("ALERTES " + getSiteConfig_(siteKey).label.toUpperCase()).setFontWeight("bold");
  sh.getRange(2, 1, 1, 5).setValues([["Date", "Type alerte", "Objet", "Détail", "ID_TECH"]]);

  if (alerts.length) {
    sh.getRange(3, 1, alerts.length, 5).setValues(alerts);
  }

  sh.autoResizeColumns(1, 5);
}

// =======================
// RUN GLOBAL
// =======================

function runAllForSite_(siteKey) {
  try {
    const result = synchroniserCalendriersVersFeuille(siteKey);
    genererDashboard_(siteKey);
    verifierConflits_(siteKey);

    SpreadsheetApp.getUi().alert(
      "Mise à jour terminée - " + result.site + "\n\n" +
      "Lignes : " + result.rows + "\n" +
      "Espaces : " + result.spaces + "\n" +
      "Ajouts : " + result.added + "\n" +
      "Modifications : " + result.updated + "\n" +
      "Suppressions : " + result.deleted
    );
  } catch (e) {
    const message =
      "Erreur pendant la mise à jour du site " + siteKey + "\n\n" +
      (e && e.message ? e.message : e);

    Logger.log(message);
    SpreadsheetApp.getUi().alert(message);
  }
}
