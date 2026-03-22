/************************************************************
 * EXPORTS VINCENNES
 ************************************************************/
function lancerExport(form) {
  const start = parserDateISO_P1_(form.dateDebut, false);
  if (!start) throw new Error("Merci de saisir une date de début valide.");

  let end;
  if (String(form.mode || "") === "25_prochains_evenements") {
    end = new Date(start.getTime());
    end.setMonth(end.getMonth() + 12);
  } else {
    end = parserDateISO_P1_(form.dateFin, true);
    if (!end) throw new Error("Merci de saisir une date de fin valide.");
    if (end < start) throw new Error("La date de fin doit être postérieure ou égale à la date de début.");
  }

  const typeExport = String(form.typeExport || "detail");
  const formatExport = String(form.formatExport || "pdf");

  const payload = construireDonneesExportVincennes_(start, end, typeExport, form);

  if (formatExport === "pdf") {
    return exporterHtmlEnPdf_(payload.title, payload.html);
  }

  if (formatExport === "excel") {
    return exporterRowsEnExcel_(payload.title, payload.rows);
  }

  throw new Error("Format d'export non reconnu.");
}

function construireDonneesExportVincennes_(start, end, typeExport, form) {
  const H = recapHelpersVincennesP1_();

  const events = listerEventsCalendar_(
    RECAP_VINCENNES_CFG_P1.calendarEvents,
    start,
    end,
    2500
  );

  events.sort((a, b) =>
    new Date(a.start.date || a.start.dateTime) - new Date(b.start.date || b.start.dateTime)
  );

  const eventsCourses = listerEventsCalendar_(
    RECAP_VINCENNES_CFG_P1.calendarCourses,
    start,
    end,
    2500
  );

  const joursCourses = construireSetJoursCourses_(eventsCourses, H);
  const eventsGroupes = regrouperEventsGlobal_(events);

  const dateDebutTxt = H.formaterDateLongue(start);
  const dateFinTxt = H.formaterDateLongue(end);

  switch (typeExport) {
    case "mensuel":
      return construireExportMensuel_(eventsGroupes, joursCourses, start, end, H, dateDebutTxt, dateFinTxt);

    case "detail":
      return construireExportDetail_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt);

    case "espace":
      return construireExportParEspace_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt);

    case "responsable":
      return construireExportParResponsable_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt);

    case "onet":
      return construireExportPrestataire_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt, "ONET");

    case "cj":
      return construireExportPrestataire_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt, "CJ");

    case "tech":
      return construireExportPrestataire_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt, "TECH");

    default:
      throw new Error("Type d'export non reconnu.");
  }
}

/************************************************************
 * BUILDERS EXPORT
 ************************************************************/
function construireExportMensuel_(eventsGroupes, joursCourses, start, end, H, dateDebutTxt, dateFinTxt) {
  const title = `Export - Récap mensuel Vincennes - ${dateDebutTxt} au ${dateFinTxt}`;
  const rows = [[
    "Mois",
    "Date début",
    "Date fin",
    "Événement",
    "Type",
    "Phase",
    "Option",
    "Jour de course",
    "Espaces"
  ]];

  eventsGroupes
    .sort((a, b) => {
      const diff = a.debutMin - b.debutMin;
      if (diff !== 0) return diff;
      return a.nom.localeCompare(b.nom, "fr", { sensitivity: "base" });
    })
    .forEach(groupe => {
      const evRef = groupe.events[0];
      const type = H.detecterType(evRef);
      const phase = libellerPhasesGroupe_(groupe);
      const isOption = groupe.isOption ? "Oui" : "";
      const toucheCourse = groupe.events.some(ev => H.eventToucheJourDeCourse(ev, joursCourses)) ? "Oui" : "";
      const mois = `${groupe.debutMin.getMonth() + 1}/${groupe.debutMin.getFullYear()}`;

      const espaces = new Set();
      groupe.events.forEach(ev => {
        if (ev.attendees) {
          ev.attendees.forEach(a => {
            if (a.resource === true) {
              espaces.add(H.extraireNomEspace(a.displayName || a.email));
            }
          });
        }
      });

      rows.push([
        mois,
        Utilities.formatDate(groupe.debutMin, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
        Utilities.formatDate(groupe.finMax, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
        groupe.nom,
        type,
        phase,
        isOption,
        toucheCourse,
        Array.from(espaces).join(", ")
      ]);
    });

  const html = construireHtmlTableSimple_(
    title,
    `Période du ${dateDebutTxt} au ${dateFinTxt}`,
    rows
  );

  return { title, rows, html };
}

function construireExportDetail_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt) {
  const title = `Export - Liste détaillée Vincennes - ${dateDebutTxt} au ${dateFinTxt}`;
  const rows = [[
    "Date début",
    "Date fin",
    "Événement",
    "Type",
    "Client",
    "Pax",
    "Réunion de courses",
    "Durée",
    "Espaces",
    "ONET",
    "CJ Sécurité",
    "Technique",
    "Autres prestataires",
    "Responsable",
    "Commentaire"
  ]];

  eventsGroupes.forEach(groupe => {
    const evRef = groupe.events[0];
    const desc = evRef.description || "";
    const type = H.detecterType(evRef);
    const client = H.extraireChamp(desc, "Client");
    const pax = H.extraireChamp(desc, "Pax");
    const responsable = H.extraireResponsable(desc);
    const reunion = groupe.events.some(ev => H.eventToucheJourDeCourse(ev, joursCourses)) ? "Oui" : "";
    const onet = H.extraireOuiNon(desc, "ONET|Prestation Nettoyage") ? "Oui" : "";
    const cj = H.extraireOuiNon(desc, "CJ SECURITE|CJ SÉCURITÉ|Prestation Sécurité") ? "Oui" : "";
    const technique = H.extraireOuiNon(desc, "Technique|Prestation Electricité") ? "Oui" : "";
    const autres = H.extraireOuiNon(desc, "Autres prestataires|Raccordage Réseau") ? "Oui" : "";
    const commentaire = H.extraireChamp(desc, "Commentaire");

    const espaces = new Set();
    groupe.events.forEach(ev => {
      if (ev.attendees) {
        ev.attendees.forEach(a => {
          if (a.resource === true) {
            espaces.add(H.extraireNomEspace(a.displayName || a.email));
          }
        });
      }
    });

    rows.push([
      Utilities.formatDate(groupe.debutMin, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      Utilities.formatDate(groupe.finMax, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      groupe.nom,
      type,
      client,
      pax,
      reunion,
      calculerDureeDepuisDates_(groupe.debutMin, groupe.finMax),
      Array.from(espaces).join(", "),
      onet,
      cj,
      technique,
      autres,
      responsable,
      commentaire
    ]);
  });

  const html = construireHtmlTableSimple_(
    title,
    `Période du ${dateDebutTxt} au ${dateFinTxt}`,
    rows
  );

  return { title, rows, html };
}

function construireExportParEspace_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt) {
  const title = `Export - Récap par espace Vincennes - ${dateDebutTxt} au ${dateFinTxt}`;
  const rows = [[
    "Espace",
    "Date début",
    "Date fin",
    "Événement",
    "Jour de course",
    "Type"
  ]];

  const recapEspaces = {};

  eventsGroupes.forEach(groupe => {
    const evRef = groupe.events[0];
    const type = H.detecterType(evRef);
    const toucheCourse = groupe.events.some(ev => H.eventToucheJourDeCourse(ev, joursCourses));

    groupe.events.forEach(ev => {
      if (ev.attendees) {
        ev.attendees.forEach(a => {
          if (a.resource === true) {
            const nom = H.extraireNomEspace(a.displayName || a.email);

            if (!recapEspaces[nom]) recapEspaces[nom] = [];
            recapEspaces[nom].push({
              event: groupe.nom,
              debut: groupe.debutMin,
              fin: groupe.finMax,
              course: toucheCourse,
              type: type
            });
          }
        });
      }
    });
  });

  Object.keys(recapEspaces)
    .sort((a, b) => a.localeCompare(b, "fr", { sensitivity: "base" }))
    .forEach(espace => {
      const itemsGroupes = regrouperOccupationsEspaceParEvent_(recapEspaces[espace]);

      itemsGroupes.forEach(item => {
        if (!item.type) {
          logDebug_(
            "construireExportParEspace_",
            "type manquant pour espace=" + espace + ", event=" + item.event
          );
        }
        rows.push([
          espace,
          Utilities.formatDate(item.debut, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
          Utilities.formatDate(item.fin, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
          item.event,
          item.course ? "Oui" : "",
          item.type || ""
        ]);
      });
    });

  logDebug_(
    "construireExportParEspace_",
    "espaces=" + Object.keys(recapEspaces).length + ", lignesExport=" + (rows.length - 1)
  );

  const html = construireHtmlTableSimple_(
    title,
    `Période du ${dateDebutTxt} au ${dateFinTxt}`,
    rows
  );

  return { title, rows, html };
}

function construireExportParResponsable_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt) {
  const title = `Export - Récap par responsable Vincennes - ${dateDebutTxt} au ${dateFinTxt}`;
  const rows = [[
    "Responsable",
    "Date début",
    "Date fin",
    "Événement",
    "Type",
    "Jour de course",
    "Client"
  ]];

  eventsGroupes.forEach(groupe => {
    const evRef = groupe.events[0];
    const desc = evRef.description || "";
    const type = H.detecterType(evRef);
    const client = H.extraireChamp(desc, "Client");
    const responsable = H.extraireResponsable(desc);
    const reunion = groupe.events.some(ev => H.eventToucheJourDeCourse(ev, joursCourses)) ? "Oui" : "";

    rows.push([
      responsable,
      Utilities.formatDate(groupe.debutMin, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      Utilities.formatDate(groupe.finMax, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      groupe.nom,
      type,
      reunion,
      client
    ]);
  });

  const header = rows[0];
  const dataRows = rows.slice(1).sort((a, b) =>
    String(a[0]).localeCompare(String(b[0]), "fr", { sensitivity: "base" })
  );
  const sortedRows = [header].concat(dataRows);

  logDebug_(
    "construireExportParResponsable_",
    "responsables=" + dataRows.length + ", premierResponsable=" + (dataRows[0] ? dataRows[0][0] : "")
  );

  const html = construireHtmlTableSimple_(
    title,
    `Période du ${dateDebutTxt} au ${dateFinTxt}`,
    sortedRows
  );

  return { title, rows: sortedRows, html };
}

function construireExportPrestataire_(eventsGroupes, joursCourses, H, dateDebutTxt, dateFinTxt, filtre) {
  const libelle =
    filtre === "ONET" ? "ONET" :
    filtre === "CJ" ? "CJ Sécurité" :
    "Technique";

  const title = `Export - ${libelle} Vincennes - ${dateDebutTxt} au ${dateFinTxt}`;
  const rows = [[
    "Date début",
    "Date fin",
    "Événement",
    "Type",
    "Client",
    "Responsable",
    "Espaces",
    "Jour de course"
  ]];

  eventsGroupes.forEach(groupe => {
    const evRef = groupe.events[0];
    const desc = evRef.description || "";
    const type = H.detecterType(evRef);
    const client = H.extraireChamp(desc, "Client");
    const responsable = H.extraireResponsable(desc);
    const reunion = groupe.events.some(ev => H.eventToucheJourDeCourse(ev, joursCourses)) ? "Oui" : "";

    let keep = false;
    if (filtre === "ONET") keep = !!H.extraireOuiNon(desc, "ONET|Prestation Nettoyage");
    if (filtre === "CJ") keep = !!H.extraireOuiNon(desc, "CJ SECURITE|CJ SÉCURITÉ|Prestation Sécurité");
    if (filtre === "TECH") keep = !!H.extraireOuiNon(desc, "Technique|Prestation Electricité");

    if (!keep) return;

    const espaces = new Set();
    groupe.events.forEach(ev => {
      if (ev.attendees) {
        ev.attendees.forEach(a => {
          if (a.resource === true) {
            espaces.add(H.extraireNomEspace(a.displayName || a.email));
          }
        });
      }
    });

    rows.push([
      Utilities.formatDate(groupe.debutMin, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      Utilities.formatDate(groupe.finMax, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      groupe.nom,
      type,
      client,
      responsable,
      Array.from(espaces).join(", "),
      reunion
    ]);
  });

  const html = construireHtmlTableSimple_(
    title,
    `Période du ${dateDebutTxt} au ${dateFinTxt}`,
    rows
  );

  return { title, rows, html };
}

/************************************************************
 * EXPORT PDF / EXCEL
 ************************************************************/
function exporterHtmlEnPdf_(title, html) {
  const fullHtml = `
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          body { font-family: Arial, sans-serif; padding: 24px; color: #111827; }
          h1 { font-size: 24px; margin: 0 0 8px 0; }
          p.meta { color: #6b7280; margin: 0 0 18px 0; }
          table { width: 100%; border-collapse: collapse; font-size: 11px; }
          th, td { border: 1px solid #d1d5db; padding: 6px 8px; vertical-align: top; }
          th { background: #f8fafc; text-align: left; }
          tr:nth-child(even) td { background: #fafafa; }
        </style>
      </head>
      <body>${html}</body>
    </html>
  `;

  const blob = Utilities.newBlob(fullHtml, "text/html", title + ".html");
  const pdfBlob = blob.getAs("application/pdf").setName(title + ".pdf");
  const file = DriveApp.createFile(pdfBlob);
  return file.getUrl();
}

function exporterRowsEnExcel_(title, rows) {
  if (!rows || !rows.length) {
    throw new Error("Aucune donnée à exporter.");
  }

  const ss = SpreadsheetApp.create(title);
  const sheet = ss.getActiveSheet();
  sheet.setName("Export");

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight("bold").setBackground("#f8fafc");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);

  const xlsxBlob = UrlFetchApp.fetch(
    "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx",
    {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    }
  ).getBlob().setName(title + ".xlsx");

  const file = DriveApp.createFile(xlsxBlob);

  DriveApp.getFileById(ss.getId()).setTrashed(true);

  return file.getUrl();
}

/************************************************************
 * HTML SIMPLE POUR PDF
 ************************************************************/
function construireHtmlTableSimple_(title, subtitle, rows) {
  let html = `<h1>${escapeHtml_(title)}</h1>`;
  html += `<p class="meta">${escapeHtml_(subtitle)}</p>`;
  html += `<table>`;

  rows.forEach((row, rowIndex) => {
    html += "<tr>";
    row.forEach(cell => {
      const tag = rowIndex === 0 ? "th" : "td";
      html += `<${tag}>${escapeHtml_(cell == null ? "" : String(cell)).replace(/\n/g, "<br>")}</${tag}>`;
    });
    html += "</tr>";
  });

  html += `</table>`;
  return html;
}

function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
