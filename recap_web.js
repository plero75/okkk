function openRecapStudio() {
  const html = HtmlService
    .createHtmlOutputFromFile("recap_studio")
    .setWidth(1440)
    .setHeight(980);

  SpreadsheetApp.getUi().showModalDialog(html, "Studio recap hippodromes");
}

function doGet(e) {
  const page = String((e && e.parameter && e.parameter.page) || "recap-studio").toLowerCase();

  if (page === "recap-studio" || page === "recap-vincennes" || page === "recap-hippodromes" || page === "studio") {
    return HtmlService
      .createHtmlOutputFromFile("recap_studio")
      .setTitle("Studio recap hippodromes")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService
    .createHtmlOutput("<h1>Page introuvable</h1><p>Utilisez <code>?page=recap-studio</code>.</p>")
    .setTitle("Page introuvable");
}

function getRecapStudioDefaults() {
  const defaults = getDialogDefaultsP1_();
  const today = new Date();
  const startOfWindow = new Date(today.getFullYear(), 0, 1);
  const endOfWindow = new Date(today.getFullYear() + 1, 11, 31);
  const recapSite = getRecapSiteConfigP1_("VINCENNES");

  return {
    dateDebut: formatRecapStudioDate_(startOfWindow),
    dateFin: formatRecapStudioDate_(endOfWindow),
    siteKey: recapSite.siteKey,
    sites: Object.keys(SITES).map(function(key) {
      return { value: key, label: SITES[key].label };
    }),
    emails: recapSite.emails || defaults.emails || [],
    emailDefaut: recapSite.emailDefaut || defaults.emailDefaut || "",
    webAppUrl: getRecapStudioUrl_()
  };
}

function getRecapStudioData(form) {
  const state = normalizeRecapStudioForm_(form);
  const context = construireContexteRecapSite_(state.siteKey, state.start, state.end);
  const preview = construireRecapSiteP1_(state.siteKey, state.start, state.end, state.destinataire);
  const viewer = getRecapStudioViewer_(state.viewerEmail, state.siteKey);

  const reports = {
    mensuel: rowsToStudioTable_(construireDonneesExportSiteDepuisContexte_(context, "mensuel", state.raw).rows),
    detail: rowsToStudioTable_(construireDonneesExportSiteDepuisContexte_(context, "detail", state.raw).rows),
    espace: rowsToStudioTable_(construireDonneesExportSiteDepuisContexte_(context, "espace", state.raw).rows),
    responsable: rowsToStudioTable_(construireDonneesExportSiteDepuisContexte_(context, "responsable", state.raw).rows),
    onet: rowsToStudioTable_(construireDonneesExportSiteDepuisContexte_(context, "onet", state.raw).rows),
    cj: rowsToStudioTable_(construireDonneesExportSiteDepuisContexte_(context, "cj", state.raw).rows),
    tech: rowsToStudioTable_(construireDonneesExportSiteDepuisContexte_(context, "tech", state.raw).rows)
  };
  const dashboard = buildRecapStudioDashboard_(context, reports, viewer);
  const profilePhotos = studioBuildProfilePhotoMap_(dashboard);

  dashboard.viewer.photoUrl = studioFindProfilePhotoUrl_(profilePhotos, dashboard.viewer.label, dashboard.viewer.email);
  dashboard.profilePhotos = profilePhotos;
  dashboard.teamBoard = studioAttachTeamPhotoUrls_(dashboard.teamBoard, profilePhotos);

  return {
    header: {
      title: "Studio recap " + context.siteLabel,
      siteKey: context.siteKey,
      siteLabel: context.siteLabel,
      siteFullLabel: context.siteFullLabel,
      periodLabel: "Du " + context.dateDebutTxt + " au " + context.dateFinTxt,
      generatedAt: Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy HH:mm"),
      webAppUrl: getRecapStudioUrl_()
    },
    dashboard: dashboard,
    profilePhotos: profilePhotos,
    stats: buildRecapStudioStats_(context, reports),
    highlights: buildRecapStudioHighlights_(reports),
    mailPreview: {
      destinataire: preview.destinataire || "",
      subject: preview.subject || "",
      html: preview.htmlBody || ""
    },
    previewHtml: preview.htmlBody,
    reports: reports
  };
}

function getRecapStudioUrl_() {
  try {
    return ScriptApp.getService().getUrl() || "";
  } catch (error) {
    return "";
  }
}

function formatRecapStudioDate_(date) {
  return Utilities.formatDate(date, TZ, "yyyy-MM-dd");
}

function isRecapStudioInternalEmail_(email) {
  return /@letrot\.com$/i.test(String(email || "").trim());
}

function normalizeRecapStudioForm_(form) {
  const raw = form || {};
  const siteKey = normalizeRecapSiteKey_(raw.siteKey);
  const recapSite = getRecapSiteConfigP1_(siteKey);
  const start = parserDateISO_P1_(raw.dateDebut, false);
  if (!start) throw new Error("Merci de renseigner une date de debut valide.");

  const end = parserDateISO_P1_(raw.dateFin, true);
  if (!end) throw new Error("Merci de renseigner une date de fin valide.");
  if (end < start) throw new Error("La date de fin doit etre posterieure ou egale a la date de debut.");

  let viewerEmail = String(raw.emailChoisi || "").trim();
  if (viewerEmail === "__custom__") {
    viewerEmail = String(raw.emailLibre || "").trim();
  }

  if (!viewerEmail || viewerEmail.indexOf("@") === -1) {
    viewerEmail = "";
  }

  let destinataire = String(raw.mailChoisi || "").trim();
  if (destinataire === "__custom__") {
    destinataire = String(raw.mailLibre || "").trim();
  }
  if (!destinataire || destinataire.indexOf("@") === -1) {
    destinataire = "";
  }
  if (!destinataire) {
    destinataire = recapSite.emailDefaut;
  }

  return {
    raw: raw,
    siteKey: siteKey,
    start: start,
    end: end,
    destinataire: destinataire,
    viewerEmail: viewerEmail
  };
}

function rowsToStudioTable_(rows) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const columns = safeRows.length ? safeRows[0].map(function(cell) {
    return String(cell == null ? "" : cell);
  }) : [];

  return {
    columns: columns,
    rows: safeRows.slice(1).map(function(row) {
      return row.map(function(cell) {
        return String(cell == null ? "" : cell);
      });
    }),
    rowCount: Math.max(safeRows.length - 1, 0)
  };
}

function buildRecapStudioStats_(context, reports) {
  const detailRows = reports.detail.rows;
  const espaceRows = reports.espace.rows;
  const responsableRows = reports.responsable.rows;

  const uniqueSpaces = new Set(espaceRows.map(function(row) {
    return row[0];
  }).filter(Boolean));

  const uniqueManagers = new Set(responsableRows.map(function(row) {
    return row[0];
  }).filter(function(value) {
    return value && value !== "Non renseigné";
  }));

  const b2bCount = detailRows.filter(function(row) {
    return row[3] === "B2B";
  }).length;
  const b2cCount = detailRows.filter(function(row) {
    return row[3] === "B2C";
  }).length;
  const courseCount = detailRows.filter(function(row) {
    return row[9] === "Oui";
  }).length;

  return [
    {
      label: "Evenements groupes",
      value: String(context.eventsGroupes.length),
      note: String(context.events.length) + " occurrences sources"
    },
    {
      label: "B2B / B2C",
      value: String(b2bCount) + " / " + String(b2cCount),
      note: "lecture metier sur la periode"
    },
    {
      label: "Espaces actifs",
      value: String(uniqueSpaces.size),
      note: "espaces mobilises"
    },
    {
      label: "Responsables",
      value: String(uniqueManagers.size),
      note: "interlocuteurs distincts"
    },
    {
      label: "Jours de course",
      value: String(context.joursCourses.size),
      note: String(courseCount) + " groupes touches"
    },
    {
      label: "Rapports",
      value: "7",
      note: "mensuel, detail, espaces, responsables, prestataires"
    }
  ];
}

function buildRecapStudioHighlights_(reports) {
  return {
    types: buildTopCountsFromRows_(reports.detail, 3, 4),
    spaces: buildTopCountsFromRows_(reports.espace, 0, 6),
    managers: buildTopCountsFromRows_(reports.responsable, 0, 6, { "Non renseigné": true }),
    months: buildTopCountsFromRows_(reports.mensuel, 0, 6)
  };
}

function buildTopCountsFromRows_(table, columnIndex, maxItems, ignoredValues) {
  const counts = {};
  const rows = table && Array.isArray(table.rows) ? table.rows : [];

  rows.forEach(function(row) {
    const value = String(row[columnIndex] || "").trim();
    if (!value) return;
    if (ignoredValues && ignoredValues[value]) return;
    counts[value] = (counts[value] || 0) + 1;
  });

  return Object.keys(counts)
    .map(function(label) {
      return { label: label, value: counts[label] };
    })
    .sort(function(left, right) {
      if (right.value !== left.value) return right.value - left.value;
      return left.label.localeCompare(right.label, "fr", { sensitivity: "base" });
    })
    .slice(0, maxItems || 5);
}

function buildRecapStudioDashboard_(context, reports, viewer) {
  const snapshots = context.eventsGroupes.map(function(groupe) {
    return buildStudioGroupSnapshot_(groupe, context);
  }).sort(sortStudioSnapshots_);
  const now = new Date();
  const threshold = context.start > now ? new Date(context.start.getTime()) : now;
  const thresholdMs = threshold.getTime();
  const upcoming = snapshots.filter(function(item) {
    return item.endMs >= thresholdMs;
  });
  const userEvents = viewer.canPersonalize ? snapshots.filter(function(item) {
    return studioGroupMatchesViewer_(item, viewer);
  }) : [];
  const userUpcoming = userEvents.filter(function(item) {
    return item.endMs >= thresholdMs;
  }).sort(sortStudioSnapshots_);

  return {
    viewer: {
      email: viewer.email,
      label: viewer.label,
      shortLabel: viewer.shortLabel,
      initials: viewer.initials,
      canPersonalize: viewer.canPersonalize,
      source: viewer.source,
      sourceLabel: viewer.sourceLabel,
      note: viewer.note
    },
    referenceLabel: "A venir a partir du " + formatStudioDayMonth_(threshold),
    globalCards: buildStudioGlobalCards_(context, snapshots, upcoming),
    userCards: buildStudioUserCards_(viewer, userEvents, userUpcoming),
    globalUpcoming: upcoming.slice(0, 8),
    userUpcoming: userUpcoming.slice(0, 8),
    teamBoard: buildStudioTeamBoard_(snapshots, thresholdMs),
    globalHighlights: buildRecapStudioHighlights_(reports),
    userHighlights: buildStudioUserHighlights_(userEvents)
  };
}

function getRecapStudioViewer_(fallbackEmail, siteKey) {
  const recapSite = getRecapSiteConfigP1_(siteKey);
  let email = "";
  let sessionEmail = "";
  let source = "none";

  try {
    sessionEmail = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  } catch (error) {
    sessionEmail = "";
  }

  const fallback = String(fallbackEmail || "").trim().toLowerCase();

  if (sessionEmail && isRecapStudioInternalEmail_(sessionEmail)) {
    email = sessionEmail;
    source = "session";
  } else if (fallback && fallback.indexOf("@") !== -1) {
    email = fallback;
    source = "selected";
  } else if (sessionEmail) {
    email = "";
    source = "team";
  }

  const label = email ? studioDisplayNameFromEmail_(email) : "Equipe " + recapSite.siteLabel;
  const localKey = studioNormalizeEmailLocalPart_(email.split("@")[0] || "");

  return {
    email: email,
    source: source,
    sourceLabel: source === "session" ? "Utilisateur connecte" : (source === "selected" ? "Profil choisi" : (source === "team" ? "Vue globale equipe" : "Vue globale")),
    label: label,
    shortLabel: studioFirstWord_(label) || label,
    initials: studioInitialsFromLabel_(label),
    canPersonalize: !!email,
    localKey: localKey,
    compactLocalKey: localKey.replace(/\s+/g, ""),
    tokens: localKey.split(" ").filter(Boolean),
    note: email
      ? (source === "session"
        ? "Le cockpit est personnalise a partir du compte Google connecte."
        : "Le cockpit est personnalise a partir du profil choisi dans le studio.")
      : (source === "team"
        ? "Aucun compte letrot.com detecte et aucun profil n'est choisi. Le studio bascule automatiquement sur la vue globale equipe."
        : "L'email du compte n'est pas expose ici. Le studio reste en vue globale tant qu'aucun email n'est disponible.")
  };
}

function buildStudioGroupSnapshot_(groupe, context) {
  const H = context.H;
  const refEvent = groupe && groupe.events && groupe.events[0] ? groupe.events[0] : {};
  const desc = String(refEvent.description || "");
  const type = String(H.detecterType(refEvent) || detecterTypeMetier_(refEvent.summary || "", desc) || "Non renseigne");
  const course = groupe.events.some(function(ev) {
    return H.eventToucheJourDeCourse(ev, context.joursCourses);
  });
  const spaces = [];
  const attendees = [];
  const managers = [];

  groupe.events.forEach(function(ev) {
    if (ev.attendees && ev.attendees.length) {
      ev.attendees.forEach(function(attendee) {
        const attendeeEmail = String((attendee && attendee.email) || "").trim().toLowerCase();
        if (attendee && attendee.resource === true) {
          const resourceName = H.extraireNomEspace(attendee.displayName || attendee.email);
          if (resourceName) spaces.push(resourceName);
        } else if (attendeeEmail) {
          attendees.push(attendeeEmail);
        }
      });
    }

    const descValue = String((ev && ev.description) || "");
    studioPushIfValue_(managers, getProjectManager_(ev));
    studioPushIfValue_(managers, H.extraireResponsable(descValue));
  });

  const ownerList = studioUniqueStrings_(managers);
  const serviceList = studioExtractGroupServices_(groupe, H);
  const uniqueSpaces = studioUniqueStrings_(spaces);
  const uniqueAttendees = studioUniqueStrings_(attendees);
  const start = new Date(groupe.debutMin);
  const end = new Date(groupe.finMax);

  return {
    key: groupe.key,
    title: String(groupe.nom || "Evenement"),
    type: type,
    tone: studioResolveTone_(type, !!groupe.isOption, course),
    phase: libellerPhasesGroupe_(groupe),
    option: !!groupe.isOption,
    course: course,
    owner: ownerList[0] || "Non renseigne",
    owners: ownerList,
    client: studioFirstNonEmpty_([
      context.H.extraireChamp(desc, "Client"),
      context.H.extraireChamp(desc, "Client final")
    ]),
    requestedOptions: getRecapRequestedOptionsLabel_(H, desc),
    spaces: uniqueSpaces,
    attendees: uniqueAttendees,
    organizerEmail: String((refEvent.organizer && refEvent.organizer.email) || "").trim().toLowerCase(),
    creatorEmail: String((refEvent.creator && refEvent.creator.email) || "").trim().toLowerCase(),
    services: serviceList,
    startMs: start.getTime(),
    endMs: end.getTime(),
    startIso: Utilities.formatDate(start, TZ, "yyyy-MM-dd'T'HH:mm:ss"),
    endIso: Utilities.formatDate(end, TZ, "yyyy-MM-dd'T'HH:mm:ss"),
    windowLabel: formatStudioWindowLabel_(start, end),
    dayLabel: Utilities.formatDate(start, TZ, "dd"),
    monthLabel: studioMonthShortLabel_(start),
    startLabel: formatStudioShortDateTime_(start),
    endLabel: formatStudioShortDateTime_(end)
  };
}

function studioExtractGroupServices_(groupe, H) {
  let hasOnet = false;
  let hasCj = false;
  let hasTech = false;

  groupe.events.forEach(function(ev) {
    const desc = String((ev && ev.description) || "");
    hasOnet = hasOnet || !!H.extraireOuiNon(desc, "ONET|Prestation Nettoyage");
    hasCj = hasCj || !!H.extraireOuiNon(desc, "CJ SECURITE|CJ SÉCURITÉ|Prestation Sécurité");
    hasTech = hasTech || !!H.extraireOuiNon(desc, "Technique|Prestation Electricité");
  });

  const labels = [];
  if (hasOnet) labels.push("ONET");
  if (hasCj) labels.push("CJ");
  if (hasTech) labels.push("Technique");
  return labels;
}

function studioGroupMatchesViewer_(snapshot, viewer) {
  if (!snapshot || !viewer || !viewer.canPersonalize) return false;
  if (snapshot.organizerEmail && snapshot.organizerEmail === viewer.email) return true;
  if (snapshot.creatorEmail && snapshot.creatorEmail === viewer.email) return true;
  if (snapshot.attendees && snapshot.attendees.indexOf(viewer.email) !== -1) return true;

  return (snapshot.owners || []).some(function(owner) {
    return studioOwnerMatchesViewer_(owner, viewer);
  });
}

function studioOwnerMatchesViewer_(owner, viewer) {
  const raw = String(owner || "").trim();
  if (!raw) return false;

  const lower = raw.toLowerCase();
  if (viewer.email && lower.indexOf(viewer.email) !== -1) return true;

  const parts = raw
    .split(/[,;/|]|(?:\set\s)/i)
    .map(function(part) { return String(part || "").trim(); })
    .filter(Boolean);

  const candidates = parts.length ? parts : [raw];

  return candidates.some(function(part) {
    const normalized = studioNormalizeIdentity_(part);
    const compact = normalized.replace(/\s+/g, "");
    if (!normalized) return false;
    if (viewer.localKey && normalized === viewer.localKey) return true;
    if (viewer.compactLocalKey && compact === viewer.compactLocalKey) return true;
    if (viewer.localKey && normalized.indexOf(viewer.localKey) !== -1) return true;

    const candidateTokens = normalized.split(" ").filter(Boolean);
    const sharedTokenCount = viewer.tokens.filter(function(token) {
      return candidateTokens.indexOf(token) !== -1;
    }).length;

    if (viewer.tokens.length >= 2 && sharedTokenCount >= 2) return true;
    if (viewer.tokens.length === 1 && candidateTokens.length === 1 && sharedTokenCount === 1) return true;
    return false;
  });
}

function buildStudioGlobalCards_(context, snapshots, upcoming) {
  const owners = new Set();
  const spaces = new Set();
  let b2bCount = 0;
  let b2cCount = 0;

  snapshots.forEach(function(item) {
    if (item.owner && item.owner !== "Non renseigne") owners.add(item.owner);
    item.spaces.forEach(function(space) { spaces.add(space); });
    if (item.type === "B2B") b2bCount++;
    if (item.type === "B2C") b2cCount++;
  });

  return [
    { icon: "calendar", tone: "blue", label: "A venir", value: String(upcoming.length), note: "dans la periode selectionnee" },
    { icon: "spark", tone: "teal", label: "Evenements groupes", value: String(snapshots.length), note: String(context.events.length) + " occurrences sources" },
    { icon: "chart", tone: "amber", label: "B2B / B2C", value: String(b2bCount) + " / " + String(b2cCount), note: "mix metier de la periode" },
    { icon: "map", tone: "pink", label: "Espaces actifs", value: String(spaces.size), note: "espaces ressources mobilises" },
    { icon: "user", tone: "navy", label: "Responsables", value: String(owners.size), note: "pilotages distincts" },
    { icon: "flag", tone: "violet", label: "Jours de course", value: String(context.joursCourses.size), note: "jalons sur la plage" }
  ];
}

function buildStudioUserCards_(viewer, userEvents, userUpcoming) {
  if (!viewer.canPersonalize) {
    return [
      { icon: "user", tone: "slate", label: "Profil perso", value: "Indispo", note: "Apps Script ne remonte pas toujours l'email du compte." },
      { icon: "mail", tone: "blue", label: "Astuce", value: "Choisis un email", note: "Le studio s'appuiera alors dessus pour la personnalisation." }
    ];
  }

  const spaces = new Set();
  let b2bCount = 0;
  let b2cCount = 0;
  let courseCount = 0;
  let optionCount = 0;

  userEvents.forEach(function(item) {
    item.spaces.forEach(function(space) { spaces.add(space); });
    if (item.type === "B2B") b2bCount++;
    if (item.type === "B2C") b2cCount++;
    if (item.course) courseCount++;
    if (item.option) optionCount++;
  });

  const nextItem = userUpcoming.length ? userUpcoming[0] : null;

  return [
    { icon: "calendar", tone: "blue", label: "Mes dates", value: String(userEvents.length), note: "sur la periode selectionnee" },
    { icon: "clock", tone: "teal", label: "Prochaine date", value: nextItem ? formatStudioDayMonth_(new Date(nextItem.startMs)) : "Aucune", note: nextItem ? nextItem.title : "rien a venir pour ce profil" },
    { icon: "map", tone: "pink", label: "Mes espaces", value: String(spaces.size), note: "espaces lies a tes dates" },
    { icon: "chart", tone: "amber", label: "B2B / B2C", value: String(b2bCount) + " / " + String(b2cCount), note: "repartition sur tes dates" },
    { icon: "flag", tone: "violet", label: "Jours de course", value: String(courseCount), note: "dates liees a une reunion" },
    { icon: "spark", tone: "navy", label: "Options", value: String(optionCount), note: "evenements sous option" }
  ];
}

function buildStudioUserHighlights_(userEvents) {
  return {
    types: buildTopCountsFromValues_(userEvents.map(function(item) { return item.type; }), 4),
    spaces: buildTopCountsFromValues_(studioFlattenValues_(userEvents.map(function(item) { return item.spaces; })), 6),
    services: buildTopCountsFromValues_(studioFlattenValues_(userEvents.map(function(item) { return item.services; })), 6)
  };
}

function buildStudioTeamBoard_(snapshots, thresholdMs) {
  const aggregated = {};

  (Array.isArray(snapshots) ? snapshots : []).forEach(function(item) {
    const owners = studioUniqueStrings_((item && item.owners && item.owners.length) ? item.owners : [item.owner]);

    owners.forEach(function(owner) {
      const label = String(owner || "").replace(/\s+/g, " ").trim();
      if (!label || label === "Non renseigne") return;

      const key = studioNormalizeIdentity_(label);
      if (!key) return;

      if (!aggregated[key]) {
        aggregated[key] = {
          label: label,
          email: studioExtractEmailFromText_(label),
          count: 0,
          nextItem: null,
          nextStartMs: Number.POSITIVE_INFINITY
        };
      }

      const bucket = aggregated[key];
      bucket.count += 1;

      if (!bucket.email) {
        bucket.email = studioExtractEmailFromText_(label);
      }

      if (item.endMs >= thresholdMs && item.startMs < bucket.nextStartMs) {
        bucket.nextItem = item;
        bucket.nextStartMs = item.startMs;
      }
    });
  });

  return Object.keys(aggregated)
    .map(function(key) {
      const bucket = aggregated[key];
      const nextItem = bucket.nextItem;
      return {
        label: bucket.label,
        email: bucket.email || "",
        initials: studioInitialsFromLabel_(bucket.label),
        count: bucket.count,
        scoreLabel: bucket.count + " date" + (bucket.count > 1 ? "s" : ""),
        nextDateLabel: nextItem ? formatStudioDayMonth_(new Date(nextItem.startMs)) : "A venir",
        subtitle: nextItem ? nextItem.title : "Aucune date a venir"
      };
    })
    .sort(function(left, right) {
      if (right.count !== left.count) return right.count - left.count;
      return String(left.label || "").localeCompare(String(right.label || ""), "fr", { sensitivity: "base" });
    })
    .slice(0, 5);
}

function buildTopCountsFromValues_(values, maxItems) {
  const counts = {};

  (Array.isArray(values) ? values : []).forEach(function(value) {
    const label = String(value || "").trim();
    if (!label) return;
    counts[label] = (counts[label] || 0) + 1;
  });

  return Object.keys(counts)
    .map(function(label) {
      return { label: label, value: counts[label] };
    })
    .sort(function(left, right) {
      if (right.value !== left.value) return right.value - left.value;
      return left.label.localeCompare(right.label, "fr", { sensitivity: "base" });
    })
    .slice(0, maxItems || 5);
}

function sortStudioSnapshots_(left, right) {
  if (left.startMs !== right.startMs) return left.startMs - right.startMs;
  return String(left.title || "").localeCompare(String(right.title || ""), "fr", { sensitivity: "base" });
}

function studioResolveTone_(type, isOption, course) {
  if (isOption) return "violet";
  if (course) return "amber";
  if (type === "B2B") return "blue";
  if (type === "B2C") return "pink";
  return "slate";
}

function studioDisplayNameFromEmail_(email) {
  const localPart = String(email || "").split("@")[0] || "";
  const parts = studioEmailLocalTokens_(localPart);
  if (!parts.length) return "Equipe";

  return parts.map(function(part) {
    const lower = part.toLowerCase();
    return lower.charAt(0).toUpperCase() + lower.slice(1);
  }).join(" ");
}

function studioInitialsFromLabel_(label) {
  const parts = String(label || "").trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "RV";
  return parts.slice(0, 2).map(function(part) {
    return part.charAt(0).toUpperCase();
  }).join("");
}

function studioFirstWord_(text) {
  return String(text || "").trim().split(/\s+/)[0] || "";
}

function studioNormalizeIdentity_(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function studioEmailLocalTokens_(value) {
  return String(value || "")
    .replace(/[._-]+/g, " ")
    .split(/\s+/)
    .map(function(part) { return String(part || "").trim(); })
    .filter(Boolean)
    .filter(function(part) {
      const normalized = studioNormalizeIdentity_(part);
      return !(normalized === "ext" || normalized === "external" || normalized === "externe");
    });
}

function studioNormalizeEmailLocalPart_(value) {
  return studioEmailLocalTokens_(value)
    .map(function(part) { return studioNormalizeIdentity_(part); })
    .filter(Boolean)
    .join(" ")
    .trim();
}

function studioUniqueStrings_(values) {
  const seen = {};
  const result = [];

  (Array.isArray(values) ? values : []).forEach(function(value) {
    const label = String(value || "").replace(/\s+/g, " ").trim();
    if (!label) return;
    const key = studioNormalizeIdentity_(label);
    if (!key || seen[key]) return;
    seen[key] = true;
    result.push(label);
  });

  return result;
}

function studioPushIfValue_(list, value) {
  const label = String(value || "").replace(/\s+/g, " ").trim();
  if (label) list.push(label);
}

function studioFirstNonEmpty_(values) {
  const items = Array.isArray(values) ? values : [];
  for (let i = 0; i < items.length; i++) {
    const label = String(items[i] || "").trim();
    if (label) return label;
  }
  return "";
}

function studioFlattenValues_(values) {
  const result = [];

  (Array.isArray(values) ? values : []).forEach(function(item) {
    if (Array.isArray(item)) {
      item.forEach(function(inner) {
        result.push(inner);
      });
      return;
    }
    result.push(item);
  });

  return result;
}

function studioBuildProfilePhotoMap_(dashboard) {
  const map = {};
  const identities = studioCollectPhotoIdentities_(dashboard);
  const selfProfile = studioFetchConnectedProfilePhoto_();

  if (selfProfile) {
    studioStoreProfilePhoto_(map, {
      label: selfProfile.name,
      email: selfProfile.email
    }, selfProfile);
  }

  identities.forEach(function(identity) {
    if (studioFindProfilePhotoUrl_(map, identity.label, identity.email)) return;
    const profile = studioLookupDirectoryProfile_(identity);
    if (profile) {
      studioStoreProfilePhoto_(map, identity, profile);
    }
  });

  return map;
}

function studioAttachTeamPhotoUrls_(teamBoard, photoMap) {
  return (Array.isArray(teamBoard) ? teamBoard : []).map(function(member) {
    const copy = {};
    Object.keys(member || {}).forEach(function(key) {
      copy[key] = member[key];
    });
    copy.photoUrl = studioFindProfilePhotoUrl_(photoMap, copy.label, copy.email);
    return copy;
  });
}

function studioCollectPhotoIdentities_(dashboard) {
  const seen = {};
  const list = [];

  function push(label, email) {
    const cleanLabel = String(label || "").replace(/\s+/g, " ").trim();
    const cleanEmail = String(email || "").trim().toLowerCase();
    const key = cleanEmail
      ? "email:" + cleanEmail
      : "label:" + studioNormalizeIdentity_(cleanLabel);
    if (!key || seen[key]) return;
    seen[key] = true;
    list.push({
      label: cleanLabel,
      email: cleanEmail
    });
  }

  const viewer = dashboard && dashboard.viewer ? dashboard.viewer : {};
  push(viewer.label, viewer.email);

  (dashboard && Array.isArray(dashboard.teamBoard) ? dashboard.teamBoard : []).forEach(function(member) {
    push(member.label, member.email);
  });

  [dashboard && dashboard.globalUpcoming, dashboard && dashboard.userUpcoming].forEach(function(items) {
    (Array.isArray(items) ? items : []).forEach(function(item) {
      const owners = studioUniqueStrings_((item && item.owners && item.owners.length) ? item.owners : [item.owner]);
      owners.forEach(function(owner) {
        push(owner, studioExtractEmailFromText_(owner));
      });
    });
  });

  return list;
}

function studioFetchConnectedProfilePhoto_() {
  if (typeof People === "undefined" || !People || !People.People || !People.People.get) return null;

  try {
    const person = People.People.get("people/me", {
      personFields: "names,emailAddresses,photos"
    });
    return studioProfileFromPerson_(person);
  } catch (error) {
    return null;
  }
}

function studioLookupDirectoryProfile_(identity) {
  const queries = studioBuildPhotoQueries_(identity);

  for (let i = 0; i < queries.length; i++) {
    const people = studioSearchDirectoryPeople_(queries[i]);
    const profile = studioSelectMatchingProfile_(people, identity);
    if (profile) return profile;
  }

  return null;
}

function studioBuildPhotoQueries_(identity) {
  const email = String(identity && identity.email || "").trim().toLowerCase();
  const label = String(identity && identity.label || "").replace(/\s+/g, " ").trim();
  const queries = [];
  const seen = {};

  function push(value) {
    const clean = String(value || "").replace(/\s+/g, " ").trim();
    if (!clean) return;
    const key = clean.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    queries.push(clean);
  }

  push(email);
  if (email) {
    push(studioDisplayNameFromEmail_(email));
  }
  push(label);
  if (label) {
    push(studioFirstWord_(label));
  }

  return queries;
}

function studioSearchDirectoryPeople_(query) {
  if (!query || typeof People === "undefined" || !People || !People.People || !People.People.searchDirectoryPeople) {
    return [];
  }

  try {
    const response = People.People.searchDirectoryPeople({
      query: query,
      readMask: "names,emailAddresses,photos",
      sources: ["DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE"]
    });
    return response && Array.isArray(response.people) ? response.people : [];
  } catch (error) {
    return [];
  }
}

function studioSelectMatchingProfile_(people, identity) {
  let best = null;
  let bestScore = 0;
  const expectedEmail = String(identity && identity.email || "").trim().toLowerCase();
  const expectedLabel = studioNormalizeIdentity_(identity && identity.label || "");
  const expectedCompact = expectedLabel.replace(/\s+/g, "");
  const expectedTokens = expectedLabel.split(" ").filter(Boolean);

  (Array.isArray(people) ? people : []).forEach(function(person) {
    const profile = studioProfileFromPerson_(person);
    if (!profile || !profile.photoUrl) return;

    const profileLabel = studioNormalizeIdentity_(profile.name);
    const profileCompact = profileLabel.replace(/\s+/g, "");
    const profileTokens = profileLabel.split(" ").filter(Boolean);
    let score = 0;

    if (expectedEmail && profile.email === expectedEmail) score += 100;
    if (expectedLabel && profileLabel === expectedLabel) score += 80;
    if (expectedCompact && profileCompact === expectedCompact) score += 70;
    if (expectedLabel && profileLabel && profileLabel.indexOf(expectedLabel) !== -1) score += 35;

    const sharedTokens = expectedTokens.filter(function(token) {
      return profileTokens.indexOf(token) !== -1;
    }).length;

    score += sharedTokens * 12;

    if (score > bestScore) {
      best = profile;
      bestScore = score;
    }
  });

  return bestScore >= 24 ? best : null;
}

function studioProfileFromPerson_(person) {
  const candidate = person || {};
  const name = studioPrimaryFieldValue_(candidate.names, "displayName");
  const email = String(studioPrimaryFieldValue_(candidate.emailAddresses, "value") || "").trim().toLowerCase();
  const photo = studioPrimaryItem_(candidate.photos);
  const photoUrl = photo && photo.url && !photo.default ? String(photo.url) : "";

  if (!name && !email && !photoUrl) return null;

  return {
    name: name || studioDisplayNameFromEmail_(email),
    email: email,
    photoUrl: photoUrl
  };
}

function studioPrimaryFieldValue_(items, fieldName) {
  const item = studioPrimaryItem_(items);
  return item && item[fieldName] ? String(item[fieldName]) : "";
}

function studioPrimaryItem_(items) {
  const list = Array.isArray(items) ? items : [];
  if (!list.length) return null;
  for (let i = 0; i < list.length; i++) {
    if (list[i] && list[i].metadata && list[i].metadata.primary) {
      return list[i];
    }
  }
  return list[0];
}

function studioStoreProfilePhoto_(map, identity, profile) {
  const url = String(profile && profile.photoUrl || "").trim();
  if (!url) return;

  [
    studioBuildPhotoKey_("email", identity && identity.email),
    studioBuildPhotoKey_("label", identity && identity.label),
    studioBuildPhotoKey_("email", profile && profile.email),
    studioBuildPhotoKey_("label", profile && profile.name)
  ].filter(Boolean).forEach(function(key) {
    map[key] = url;
  });
}

function studioFindProfilePhotoUrl_(photoMap, label, email) {
  const emailKey = studioBuildPhotoKey_("email", email);
  if (emailKey && photoMap && photoMap[emailKey]) return photoMap[emailKey];

  const labelKey = studioBuildPhotoKey_("label", label);
  if (labelKey && photoMap && photoMap[labelKey]) return photoMap[labelKey];

  return "";
}

function studioBuildPhotoKey_(type, value) {
  const normalized = type === "email"
    ? String(value || "").trim().toLowerCase()
    : studioNormalizeIdentity_(value);
  return normalized ? type + ":" + normalized : "";
}

function studioExtractEmailFromText_(value) {
  const match = String(value || "").toLowerCase().match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/);
  return match ? match[0] : "";
}

function formatStudioShortDateTime_(date) {
  return Utilities.formatDate(date, TZ, "dd/MM HH:mm");
}

function formatStudioDayMonth_(date) {
  return Utilities.formatDate(date, TZ, "dd/MM/yyyy");
}

function formatStudioWindowLabel_(start, end) {
  const sameDay = Utilities.formatDate(start, TZ, "yyyyMMdd") === Utilities.formatDate(end, TZ, "yyyyMMdd");
  if (sameDay) {
    return Utilities.formatDate(start, TZ, "dd/MM/yyyy HH:mm") + " - " + Utilities.formatDate(end, TZ, "HH:mm");
  }
  return Utilities.formatDate(start, TZ, "dd/MM/yyyy HH:mm") + " -> " + Utilities.formatDate(end, TZ, "dd/MM/yyyy HH:mm");
}

function studioMonthShortLabel_(date) {
  const months = ["Jan", "Fev", "Mar", "Avr", "Mai", "Jun", "Jul", "Aou", "Sep", "Oct", "Nov", "Dec"];
  return months[date.getMonth()] || "";
}
