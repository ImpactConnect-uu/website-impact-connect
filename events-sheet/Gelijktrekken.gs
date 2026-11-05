/* ═══ GELIJKTREKKEN: Events ⇄ categorietabbladen ═════════════════════
   Hoort bij de Google Sheet "LINKED TO BOT Events Database", niet bij de
   formulier-backend. Je plakt dit in die sheet via Extensies → Apps Script,
   in een eigen bestand. Een push zet het nergens neer.

   Wat het doet:
     1. Pas je iets aan in Events, dan krijgt elk categorietabblad waar dat
        event onder valt dezelfde aanpassing.
     2. Pas je iets aan in een categorietabblad, dan gaat dat eerst naar
        Events en van daaruit naar de andere tabbladen.
     3. Verander je de categorie, dan verhuist het event naar het juiste
        tabblad. Een nieuw event in een categorietabblad komt in Events.
     4. Elk uur, en via het menu Impact Connect → Alles gelijktrekken,
        worden alle categorietabbladen opnieuw opgebouwd uit Events. Dat
        vangt op wat het script niet ziet: een verwijderde rij, of iets
        wat via een koppeling (de bot, Claude) is veranderd.

     5. Elke rij krijgt dezelfde opmaak: afwisselend lichtgroen en wit,
        randen, Arial 9, tekst die doorloopt, de naam vet. De kleuren zijn
        "afwisselende kleuren" van Google zelf, dus een nieuwe rij kleurt
        vanzelf mee. Plak je iets met een eigen opmaak, dan zet het script
        die meteen terug. Het filter loopt tot onderaan het tabblad, zodat
        ook nieuwe events erin vallen.

   Bij twijfel wint Events. Verwijderen doe je dus in Events; een rij die je
   alleen uit een categorietabblad haalt komt bij het gelijktrekken terug.

   De kolommen gaan op naam van de kop, niet op plek. Een kop hernoemen in
   een categorietabblad maakt die kolom dus leeg; zet hem terug en draai
   Alles gelijktrekken.

   Alle namen beginnen met SYNC_ of sync, zodat dit niet botst met andere
   scripts in dezelfde sheet.

   Eenmalig aanzetten: kies bovenin syncInstalleren, klik Uitvoeren en
   daarna Toestaan.
   ═══════════════════════════════════════════════════════════════════ */

const SYNC_HOOFDTAB = 'Events';

/* Welk tabblad bij welke categorie hoort. "woord" is waarop we zoeken in
   de kolom Category, net als het tabblad Overzicht doet; "naam" is wat we
   invullen als iemand een nieuw event in dat tabblad typt. */
const SYNC_TABS = [
  { tab: 'General',               woord: 'general',    naam: 'General' },
  { tab: 'Business & Innovation', woord: 'business',   naam: 'Business and innovation' },
  { tab: 'Ecology & Food',        woord: 'ecology',    naam: 'Ecology and food' },
  { tab: 'Energy',                woord: 'energy',     naam: 'Energy' },
  { tab: 'Governance',            woord: 'governance', naam: 'Governance' },
  { tab: 'Cities',                woord: 'cities',     naam: 'Cities' },
];

const SYNC_KOP_NAAM      = 'Event';
const SYNC_KOP_CATEGORIE = 'Category';
const SYNC_KOP_DATUM     = 'Date';

const SYNC_FUNCTIES = ['syncBijBewerken', 'syncBijOpenen', 'syncGelijktrekken'];

/* De opmaak van de rijen onder de kop, overgenomen van wat er al stond.
   "vooruit" is hoeveel lege rijen onder het laatste event alvast opgemaakt
   worden, zodat wat je daar typt of plakt er meteen goed uitziet. */
const SYNC_OPMAAK = {
  groen: '#f1f8e9',
  wit: '#ffffff',
  rand: '#bdbdbd',
  letter: 'Arial',
  grootte: 9,
  vooruit: 50,
};

/* ---- eenmalig: de drie triggers zetten ----------------------------- */
function syncInstalleren() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (SYNC_FUNCTIES.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncBijBewerken').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('syncBijOpenen').forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger('syncGelijktrekken').timeBased().everyHours(1).create();
  const uitkomst = syncGelijktrekken();
  console.log('Aangezet. ' + uitkomst);
}

/* Zet alles weer uit. De tabbladen blijven zoals ze zijn. */
function syncUitzetten() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (SYNC_FUNCTIES.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  console.log('Uitgezet. Er wordt niets meer vanzelf gelijkgetrokken.');
}

/* ---- het menu ------------------------------------------------------ */
function syncBijOpenen() {
  SpreadsheetApp.getUi().createMenu('Impact Connect')
    .addItem('Alles gelijktrekken', 'syncNuMetMelding')
    .addToUi();
}

function syncNuMetMelding() {
  SpreadsheetApp.getActive().toast(syncGelijktrekken() + ' Opmaak bijgewerkt.', 'Gelijktrekken', 8);
}

/* ---- bij elke wijziging -------------------------------------------- */
function syncBijBewerken(e) {
  if (!e || !e.range) return;
  const blad = e.range.getSheet();
  const def = syncTabDef(blad.getName());
  if (blad.getName() !== SYNC_HOOFDTAB && !def) return;
  if (e.range.getLastRow() < 2) return;              // alleen de kopregel

  // Eerst de opmaak van wat er net getypt of geplakt is. Die hangt niet af
  // van de plek van de rij, dus hij klopt ook als de rij straks verschuift.
  const vanRij = Math.max(2, e.range.getRow());
  syncOpmaakRijen(blad, vanRij, e.range.getLastRow() - vanRij + 1);

  const slot = LockService.getDocumentLock();
  if (!slot.tryLock(25000)) return;                  // het uur-rondje haalt het in
  try {
    if (def && !syncNaarEvents(e, blad, def)) return;
    syncGelijktrekkenBinnenSlot();
  } finally {
    slot.releaseLock();
  }
}

/* Zet een wijziging uit een categorietabblad over naar Events. Geeft false
   terug als er nog niets over te zetten valt: dan typt iemand een nieuwe
   rij en staat de naam er nog niet. Dan laten we het tabblad met rust,
   anders verdwijnt wat hij al had getypt. */
function syncNaarEvents(e, blad, def) {
  const hoofd = blad.getParent().getSheetByName(SYNC_HOOFDTAB);
  const hKoppen = syncKoppen(hoofd);
  const tKoppen = syncKoppen(blad);
  const hNaam = hKoppen.indexOf(SYNC_KOP_NAAM);
  const tNaam = tKoppen.indexOf(SYNC_KOP_NAAM);
  if (hNaam < 0 || tNaam < 0) return false;

  const r = e.range;
  const eersteRij = Math.max(2, r.getRow());
  const laatsteRij = r.getLastRow();
  const eersteKol = r.getColumn();
  const laatsteKol = Math.min(r.getLastColumn(), tKoppen.length);
  const tabRijen = blad.getRange(eersteRij, 1, laatsteRij - eersteRij + 1, tKoppen.length).getValues();
  const hoofdWaarden = hoofd.getDataRange().getValues();

  // Eén cel in de kolom Event veranderd: dan is de oude naam het event
  // waar het om gaat. Zo wordt een hernoemd event geen tweede event.
  const alleenNaamCel = r.getNumRows() === 1 && r.getNumColumns() === 1
    && eersteKol - 1 === tNaam && syncTekst(e.oldValue) !== '';

  let iets = false;
  tabRijen.forEach(function (rij) {
    const naam = syncTekst(rij[tNaam]);
    if (!naam) {
      // Een bestaande naam gewist: niet uit Events halen, maar het tabblad
      // opnieuw opbouwen, dan staat hij er meteen weer.
      if (alleenNaamCel) iets = true;
      return;
    }
    iets = true;
    const sleutel = alleenNaamCel ? syncTekst(e.oldValue) : naam;
    const doel = syncZoekRij(hoofdWaarden, hNaam, sleutel);

    if (doel > 0) {
      for (let k = eersteKol; k <= laatsteKol; k++) {
        const hk = hKoppen.indexOf(tKoppen[k - 1]);
        if (!tKoppen[k - 1] || hk < 0) continue;
        hoofd.getRange(doel + 1, hk + 1).setValue(rij[k - 1]);
        hoofdWaarden[doel][hk] = rij[k - 1];
      }
    } else {
      // Een nieuw event. Zonder deze categorie zou het meteen weer uit
      // het tabblad verdwijnen waarin hij net getypt is.
      const nieuw = syncNieuweRij(hKoppen, tKoppen, rij, def);
      hoofd.appendRow(nieuw);
      hoofdWaarden.push(nieuw);
    }
  });
  return iets;
}

/* ---- alles opnieuw opbouwen uit Events ----------------------------- */
function syncGelijktrekken() {
  const slot = LockService.getDocumentLock();
  if (!slot.tryLock(25000)) return 'Even bezig met iets anders; probeer het zo nog eens.';
  let uitkomst;
  try {
    uitkomst = syncGelijktrekkenBinnenSlot();
  } finally {
    slot.releaseLock();
  }
  // De opmaak na het slot: die raakt geen inhoud, en zo hoeft niemand die
  // net iets typt erop te wachten.
  syncOpmaakAlles();
  return uitkomst;
}

function syncGelijktrekkenBinnenSlot() {
  const ss = SpreadsheetApp.getActive();
  const hoofd = ss.getSheetByName(SYNC_HOOFDTAB);
  if (!hoofd) throw new Error('Geen tabblad "' + SYNC_HOOFDTAB + '". Is het hernoemd?');
  const waarden = hoofd.getDataRange().getValues();
  const events = { koppen: waarden[0].map(syncTekst), rijen: waarden.slice(1) };

  // Liever niets doen dan alle tabbladen leeggooien omdat Events even
  // onleesbaar is.
  const hNaam = events.koppen.indexOf(SYNC_KOP_NAAM);
  if (hNaam < 0 || events.koppen.indexOf(SYNC_KOP_CATEGORIE) < 0) {
    throw new Error('Events mist de kolom "' + SYNC_KOP_NAAM + '" of "' + SYNC_KOP_CATEGORIE + '". Niets veranderd.');
  }
  const aantal = events.rijen.filter(function (r) { return syncTekst(r[hNaam]); }).length;
  if (!aantal) throw new Error('Events lijkt leeg. Niets veranderd.');

  const veranderd = [];
  SYNC_TABS.forEach(function (def) {
    const blad = ss.getSheetByName(def.tab);
    if (!blad) return;
    if (syncVulTab(blad, events, def)) veranderd.push(def.tab);
  });
  return aantal + ' events in Events. '
    + (veranderd.length ? 'Bijgewerkt: ' + veranderd.join(', ') + '.' : 'Alles stond al gelijk.');
}

/* Schrijft alleen als er echt iets anders is, zodat een tabblad niet bij
   elke wijziging ergens anders opnieuw wordt neergezet. */
function syncVulTab(blad, events, def) {
  const koppen = syncKoppen(blad);
  if (koppen.indexOf(SYNC_KOP_NAAM) < 0) return false;
  const breedte = koppen.length;
  const nieuw = syncRijenVoorTab(events, koppen, def);
  const oudAantal = Math.max(blad.getLastRow() - 1, 0);
  const oud = oudAantal ? blad.getRange(2, 1, oudAantal, breedte).getValues() : [];
  if (JSON.stringify(oud) === JSON.stringify(nieuw) && syncAlleenWaarden(blad, oudAantal, breedte)) return false;

  // clearContent laat de opmaak en kleuren staan
  if (oudAantal) blad.getRange(2, 1, oudAantal, breedte).clearContent();
  if (nieuw.length) blad.getRange(2, 1, nieuw.length, breedte).setValues(nieuw);
  return true;
}

/* Staat er nog een formule in het tabblad (zoals de oude XLOOKUP in Good to
   send?), dan moet die eruit, ook als de uitkomst toevallig gelijk is. */
function syncAlleenWaarden(blad, aantal, breedte) {
  if (!aantal) return true;
  return !blad.getRange(2, 1, aantal, breedte).getFormulas()
    .some(function (r) { return r.some(function (f) { return f !== ''; }); });
}

/* ---- de opmaak ------------------------------------------------------- */

/* Events en alle categorietabbladen. */
function syncOpmaakAlles() {
  const ss = SpreadsheetApp.getActive();
  [SYNC_HOOFDTAB].concat(SYNC_TABS.map(function (d) { return d.tab; })).forEach(function (naam) {
    const blad = ss.getSheetByName(naam);
    if (blad) syncOpmaak(blad);
  });
}

/* Eén tabblad: de rijen met events plus een stuk lege rijen eronder, de
   afwisselende kleuren tot onderaan, en het filter tot onderaan. */
function syncOpmaak(blad) {
  const breedte = syncKoppen(blad).length;
  if (!breedte || blad.getMaxRows() < 2) return;
  const laatste = Math.min(blad.getLastRow() + SYNC_OPMAAK.vooruit, blad.getMaxRows());
  syncOpmaakRijen(blad, 2, laatste - 1, breedte);

  // Lege rijen alvast als tekst, anders maakt Google van 21-11-2026 een
  // datum en leest de bot er iets anders in. Gevulde cellen blijven zoals
  // ze zijn: daar staat de tekst al goed.
  const leegVanaf = Math.max(blad.getLastRow() + 1, 2);
  if (leegVanaf <= laatste) blad.getRange(leegVanaf, 1, laatste - leegVanaf + 1, breedte).setNumberFormat('@');

  syncBanding(blad, breedte);
  syncFilter(blad, breedte);
}

/* Lettertype, randen en uitlijning. De achtergrond gaat eraf, want de
   kleuren komen van de afwisselende kleuren; een eigen kleur op een cel
   zou daar overheen liggen. */
function syncOpmaakRijen(blad, vanRij, aantal, breedte) {
  breedte = breedte || syncKoppen(blad).length;
  if (!breedte || aantal < 1) return;
  blad.getRange(vanRij, 1, aantal, breedte)
    .setBackground(null)
    .setFontFamily(SYNC_OPMAAK.letter)
    .setFontSize(SYNC_OPMAAK.grootte)
    .setFontWeight('normal')
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, true, true, SYNC_OPMAAK.rand, SpreadsheetApp.BorderStyle.SOLID);
  blad.getRange(vanRij, 1, aantal, 1).setFontWeight('bold');
}

/* Afwisselend lichtgroen en wit, van rij 2 tot onderaan. Staat het al zo,
   dan blijft het staan. */
function syncBanding(blad, breedte) {
  const hoogte = blad.getMaxRows() - 1;
  const bestaand = blad.getBandings();
  if (bestaand.length === 1) {
    const r = bestaand[0].getRange();
    if (r.getRow() === 2 && r.getColumn() === 1 && r.getNumRows() === hoogte
        && r.getNumColumns() === breedte) return;
  }
  bestaand.forEach(function (b) { b.remove(); });
  blad.getRange(2, 1, hoogte, breedte)
    .applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, false, false)
    .setFirstRowColor(SYNC_OPMAAK.groen)
    .setSecondRowColor(SYNC_OPMAAK.wit);
}

/* Het filter liep tot rij 54, dus een nieuw event viel erbuiten. Nu tot
   onderaan, met dezelfde instellingen (bij Events: No verbergen). */
function syncFilter(blad, breedte) {
  const filter = blad.getFilter();
  if (!filter) return;
  const r = filter.getRange();
  const hoogte = blad.getMaxRows();
  const kolommen = Math.max(breedte, r.getLastColumn());
  if (r.getRow() === 1 && r.getColumn() === 1 && r.getNumRows() === hoogte
      && r.getNumColumns() === kolommen) return;
  const criteria = [];
  for (let c = r.getColumn(); c <= r.getLastColumn(); c++) {
    const k = filter.getColumnFilterCriteria(c);
    if (k) criteria.push([c, k]);
  }
  filter.remove();
  const nieuw = blad.getRange(1, 1, hoogte, kolommen).createFilter();
  criteria.forEach(function (p) { nieuw.setColumnFilterCriteria(p[0], p[1]); });
}

/* ═══ Hieronder alleen rekenwerk, zonder de sheet aan te raken ═════════ */

function syncTabDef(naam) {
  for (let i = 0; i < SYNC_TABS.length; i++) if (SYNC_TABS[i].tab === naam) return SYNC_TABS[i];
  return null;
}

function syncTekst(w) {
  return w === null || w === undefined ? '' : String(w).trim();
}

function syncKoppen(blad) {
  const breedte = blad.getLastColumn();
  if (!breedte) return [];
  const koppen = blad.getRange(1, 1, 1, breedte).getValues()[0].map(syncTekst);
  while (koppen.length && !koppen[koppen.length - 1]) koppen.pop();
  return koppen;
}

function syncHoortBij(categorie, def) {
  return syncTekst(categorie).toLowerCase().indexOf(def.woord) !== -1;
}

/* Rij-index in waarden (1 of hoger), of -1. Hoofdletters en spaties aan
   het begin of eind tellen niet mee. */
function syncZoekRij(waarden, kolom, naam) {
  const zoek = syncTekst(naam).toLowerCase();
  for (let i = 1; i < waarden.length; i++) {
    if (syncTekst(waarden[i][kolom]).toLowerCase() === zoek) return i;
  }
  return -1;
}

function syncNieuweRij(hKoppen, tKoppen, rij, def) {
  const nieuw = hKoppen.map(function (kop) {
    const tk = tKoppen.indexOf(kop);
    return kop && tk >= 0 ? rij[tk] : '';
  });
  const hc = hKoppen.indexOf(SYNC_KOP_CATEGORIE);
  if (hc >= 0 && !syncHoortBij(nieuw[hc], def)) {
    nieuw[hc] = syncTekst(nieuw[hc]) ? syncTekst(nieuw[hc]) + ', ' + def.naam : def.naam;
  }
  return nieuw;
}

/* De rijen die in een categorietabblad horen, met de kolommen van dat
   tabblad, gesorteerd op de eerste datum. Zonder leesbare datum achteraan,
   in de volgorde van Events. */
function syncRijenVoorTab(events, tabKoppen, def) {
  const hNaam = events.koppen.indexOf(SYNC_KOP_NAAM);
  const hCat = events.koppen.indexOf(SYNC_KOP_CATEGORIE);
  const hDatum = events.koppen.indexOf(SYNC_KOP_DATUM);
  const passend = [];
  events.rijen.forEach(function (rij, i) {
    if (!syncTekst(rij[hNaam]) || !syncHoortBij(rij[hCat], def)) return;
    passend.push({ i: i, datum: hDatum < 0 ? Infinity : syncDatum(rij[hDatum]), rij: rij });
  });
  passend.sort(function (a, b) { return a.datum - b.datum || a.i - b.i; });
  return passend.map(function (p) {
    return tabKoppen.map(function (kop) {
      const hk = events.koppen.indexOf(kop);
      return kop && hk >= 0 ? p.rij[hk] : '';
    });
  });
}

const SYNC_MAANDEN = {
  jan: 1, januari: 1, january: 1, feb: 2, februari: 2, february: 2, mrt: 3, mar: 3,
  maart: 3, march: 3, apr: 4, april: 4, mei: 5, may: 5, jun: 6, juni: 6, june: 6,
  jul: 7, juli: 7, july: 7, aug: 8, augustus: 8, august: 8, sep: 9, sept: 9,
  september: 9, okt: 10, oct: 10, oktober: 10, october: 10, nov: 11, november: 11,
  dec: 12, december: 12,
};

/* De eerste datum in de cel, als getal om op te sorteren. Kent 12-02-2027,
   6-10-2026 to 8-10-2026, 21 November 2026 en 1 mei 2027. */
function syncDatum(w) {
  if (w instanceof Date) return w.getTime();
  const t = syncTekst(w).toLowerCase();
  let m = t.match(/(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m) return Date.UTC(+m[3], +m[2] - 1, +m[1]);
  m = t.match(/(\d{1,2})\s+([a-z]+)\.?\s+(\d{4})/);
  if (m && SYNC_MAANDEN[m[2]]) return Date.UTC(+m[3], SYNC_MAANDEN[m[2]] - 1, +m[1]);
  return Infinity;
}
