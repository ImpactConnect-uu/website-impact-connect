/* ===================================================================
   Alumniverhalen: "what is the job really like?"
   ===================================================================
   Alumni krijgen één persoonlijke mail met wat we al van ze weten (baan,
   eerdere rol, studie) en een link naar verhaal.html. Daar beantwoorden ze
   zoveel van de zes vragen als ze willen. De antwoorden komen in het
   tabblad "Alumni verhalen" van deze sheet.

   Op de site (paths.html) komt een verhaal pas als jullie twee dingen doen:
     1. in de kolom "Op de site" yes zetten, en
     2. in "Op de site als" een anonieme omschrijving typen, bijvoorbeeld
        "Data analyst at a large consultancy". Zonder die omschrijving blijft
        het verhaal weg, zodat er nooit per ongeluk een werkgever op de site
        staat.
   En natuurlijk alleen als de alumnus zelf het vinkje voor publiceren zette.

   Wat de site te zien krijgt bepaalt verhalenOpenbaar() hieronder: geen
   naam, geen mailadres, geen baan zoals de alumnus die zelf schreef.

   De mail versturen gaat in drie stappen, in de Apps Script-editor:
     testVerhaalMail()       alleen naar het testprofiel in de alumnisheet
     toonVerhaalWachtrij()   wie hem zou krijgen, zonder iets te versturen
     stuurVerhaalMails()     echt naar iedereen, meteen. Werkt pas als
                             VERHAAL_MAILS_AAN hieronder op true staat.
   Of in plaats van die laatste: planVerhaalMailsMorgen() verstuurt de ronde
   morgen om 09:30, één keer; annuleerVerhaalMails() zet dat weer uit.
   Wie de mail al kreeg, krijgt hem nooit een tweede keer.
   =================================================================== */

/* Staat bewust uit. Pas op true zetten (en pushen) als jullie de mail
   hebben getest en akkoord zijn. Dan pas doet stuurVerhaalMails() iets. */
const VERHAAL_MAILS_AAN = true;          // aangezet op 25 november 2026

const VERHAAL_TABBLAD = 'Alumni verhalen';
const VERHAAL_LINK_DAGEN = 180;          // zo lang werkt de link in de mail

/* De zes vragen: id zoals de site hem meestuurt, en de kolom in de sheet.
   De vragen zelf staan op de site (ALUMNI_VRAGEN in script.js). */
const VERHAAL_VRAGEN = [
  { id: 'dag',     kolom: 'A typical day' },
  { id: 'pad',     kolom: 'How I got here' },
  { id: 'skills',  kolom: 'Skills' },
  { id: 'leuk',    kolom: 'You must enjoy' },
  { id: 'niet',    kolom: 'Not for you if' },
  { id: 'tips',    kolom: 'Tips' },
];

/* Het script schrijft op kolomnaam: verslepen mag, hernoemen niet. De laatste
   drie vullen jullie zelf in; het script laat ze daarna met rust. */
const VERHAAL_KOLOMMEN = ['Ingevuld op', 'Alumnus id', 'Naam', 'E-mail',
  'Huidige baan', 'Eerdere rol', 'Studie']
  .concat(VERHAAL_VRAGEN.map(function (v) { return v.kolom; }))
  .concat(['Mag op de site', 'Op de site', 'Op de site als', 'Richting']);

/* ---- het tabblad ------------------------------------------------- */
function verhaalBlad() {
  const bestand = SpreadsheetApp.getActiveSpreadsheet();
  if (!bestand) return null;
  let blad = bestand.getSheetByName(VERHAAL_TABBLAD);
  if (!blad) {
    blad = bestand.insertSheet(VERHAAL_TABBLAD);
    blad.appendRow(VERHAAL_KOLOMMEN);
    blad.setFrozenRows(1);
    blad.getRange(1, 1, 1, VERHAAL_KOLOMMEN.length).setFontWeight('bold');
  }
  return blad;
}

function verhaalKop(blad) {
  return blad.getRange(1, 1, 1, Math.max(blad.getLastColumn(), 1))
    .getValues()[0].map(function (v) { return String(v).trim(); });
}

/* ---- wie hoort bij een link -------------------------------------- */
function verhaalLink(alumnus) {
  return SITE_ADRES + '/verhaal.html?t='
    + encodeURIComponent(maakToken('verhaal', alumnus.id, '', VERHAAL_LINK_DAGEN));
}

function alumnusVanToken(token) {
  const p = leesToken(token, 'verhaal');
  if (!p) return null;
  return leesAlumni().filter(function (a) { return a.id === p.ref; })[0] || null;
}

function studieVan(a) {
  return [a.bsc, a.msc].filter(String).join(', ');
}

/* Een eerder ingevuld verhaal van deze alumnus, zodat een tweede keer de
   vakken al gevuld zijn en de eerste regel wordt overschreven. */
function verhaalRij(blad, id) {
  if (!blad || blad.getLastRow() < 2) return null;
  const kop = verhaalKop(blad);
  const k = kop.indexOf('Alumnus id');
  if (k === -1) return null;
  const waarden = blad.getRange(2, 1, blad.getLastRow() - 1, kop.length).getValues();
  for (let i = 0; i < waarden.length; i++) {
    if (String(waarden[i][k]).trim() === id) return { nummer: i + 2, kop: kop, rij: waarden[i] };
  }
  return null;
}

/* ---- wat verhaal.html ophaalt -------------------------------------
   Alleen voor wie de persoonlijke link heeft. Gaat terug: voornaam en wat we
   al van de alumnus weten, plus wat hij eventueel al eerder invulde. */
function verhaalOpzet(token) {
  const a = alumnusVanToken(token);
  if (!a) return { ok: false, reden: 'link' };
  const uit = {
    ok: true,
    voornaam: toonNaam(a.volledigeNaam),
    werk: a.werk,
    eerder: a.eerder,
    studie: studieVan(a),
    antwoorden: {},
    publiceren: false,                 // het vinkje staat uit tot de alumnus het aanzet
  };
  const oud = verhaalRij(verhaalBlad(), a.id);
  if (oud) {
    const cel = function (naam) {
      const i = oud.kop.indexOf(naam);
      return i === -1 ? '' : String(oud.rij[i] || '').trim();
    };
    if (cel('Huidige baan')) uit.werk = cel('Huidige baan');
    if (cel('Eerdere rol')) uit.eerder = cel('Eerdere rol');
    if (cel('Studie')) uit.studie = cel('Studie');
    VERHAAL_VRAGEN.forEach(function (v) { uit.antwoorden[v.id] = cel(v.kolom); });
    uit.publiceren = cel('Mag op de site').toLowerCase() === 'yes';
  }
  return uit;
}

/* ---- opslaan ------------------------------------------------------ */
function verhaalOpslaan(d) {
  const a = alumnusVanToken(d.token);
  if (!a) return antwoord({ ok: false, reden: 'link' });

  const antwoorden = d.antwoorden || {};
  const waarden = {
    'Ingevuld op': new Date(),
    'Alumnus id': a.id,
    'Naam': a.volledigeNaam,
    'E-mail': a.mail,
    'Huidige baan': tekstOf(d.werk, a.werk),
    'Eerdere rol': tekstOf(d.eerder, a.eerder),
    'Studie': tekstOf(d.studie, studieVan(a)),
    'Mag op de site': (d.publiceren === true || d.publiceren === 'true') ? 'yes' : 'no',
  };
  let ingevuld = 0;
  VERHAAL_VRAGEN.forEach(function (v) {
    waarden[v.kolom] = tekstOf(antwoorden[v.id], '');
    if (waarden[v.kolom]) ingevuld++;
  });
  if (!ingevuld) return antwoord({ ok: false, reden: 'leeg' });

  // Een voorstel voor de twee kolommen die jullie anders zelf invullen. Alleen
  // in een lege cel: wat jullie er al in typten blijft staan.
  const voorstel = {
    'Op de site als': rolZonderWerkgever(waarden['Huidige baan']),
    'Richting': richtingVan(waarden['Huidige baan'], a.interessesRuw),
  };

  let gelukt = false;
  let tweedeKeer = false;
  metSlot(function () {
    const blad = verhaalBlad();
    if (!blad) return;
    const oud = verhaalRij(blad, a.id);
    if (oud) {
      // Overschrijven, maar alleen de kolommen van de alumnus zelf. Wat jullie
      // invulden (Op de site, Op de site als, Richting) blijft staan.
      tweedeKeer = true;
      Object.keys(waarden).forEach(function (naam) {
        const i = oud.kop.indexOf(naam);
        if (i !== -1) blad.getRange(oud.nummer, i + 1).setValue(waarden[naam]);
      });
      vulLegeCellen(blad, oud.nummer, oud.kop, voorstel);
    } else {
      const kop = verhaalKop(blad);
      const alles = Object.assign({}, voorstel, waarden);
      blad.appendRow(kop.map(function (naam) {
        return alles[naam] === undefined ? '' : alles[naam];
      }));
    }
    gelukt = true;
  });
  if (!gelukt) return antwoord({ ok: false, reden: 'sheet' });

  meldVerhaal(waarden, ingevuld, tweedeKeer, voorstel);
  return antwoord({ ok: true });
}

/* ---- het voorstel voor "Op de site als" en "Richting" ---------------
   In de mail aan alumni beloven we hun werkgever niet te noemen. Daarom de
   rol zonder alles na "at", "bij", "for" of "@": "Research scientist at TNO"
   wordt "Research scientist". Staat er maar één woord, dan is dat meestal een
   bedrijfsnaam ("EnergieStudent") en blijft de cel leeg. Lees hem dus altijd
   na voordat je "Op de site" op yes zet. */
function rolZonderWerkgever(werk) {
  const ruw = String(werk || '').replace(/\s+/g, ' ').trim();
  if (!ruw) return '';
  const delen = ruw.split(/\s+(?:at|bij|for|voor)\s+|\s*@\s*|\s+[|–-]\s+/i);
  const rol = delen[0].replace(/[,;:]+$/, '').trim();
  // Eén los woord: een herkenbare functie ("Teacher", "Consultant") houden we,
  // een woord met een hoofdletter middenin ("EnergieStudent") is een merknaam.
  if (delen.length === 1 && rol.indexOf(' ') === -1) {
    const isFunctie = groepeerThemas(rol).length && !/[a-z][A-Z]/.test(rol);
    if (!isFunctie) return '';
  }
  return rol.charAt(0).toUpperCase() + rol.slice(1);
}

/* De richting met dezelfde thema's als de alumnipagina (THEMAS in Loket.gs).
   Eerst op de baan, anders op de interessegebieden. De volgorde hieronder
   bepaalt wat wint als er meer passen: het soort werk gaat voor het onderwerp.
   "Choosing a master's" en "Job hunting" gaan niet over een baan en tellen
   hier niet mee. */
const RICHTING_VOORRANG = ['Consultancy', 'IT & data', 'Engineering', 'Finance', 'Law',
  'Health', 'Education', 'Research & PhD', 'Policy & government', 'Entrepreneurship',
  'Marketing & communications', 'Sustainability', 'Working abroad'];

function richtingVan(werk, interesses) {
  const kies = function (tekst) {
    // "Utrecht" bevat "recht" en zou anders altijd Law opleveren.
    const gevonden = groepeerThemas(String(tekst || '').replace(/utrecht/gi, ' '));
    return RICHTING_VOORRANG.filter(function (r) { return gevonden.indexOf(r) !== -1; })[0] || '';
  };
  // Eerst de rol zelf, dan de hele baan met werkgever, dan de interesses.
  return kies(rolZonderWerkgever(werk)) || kies(werk) || kies(interesses);
}

function vulLegeCellen(blad, nummer, kop, voorstel) {
  Object.keys(voorstel).forEach(function (naam) {
    const i = kop.indexOf(naam);
    if (i === -1 || !voorstel[naam]) return;
    const cel = blad.getRange(nummer, i + 1);
    if (!String(cel.getValue()).trim()) cel.setValue(voorstel[naam]);
  });
}

/* Voor de verhalen die al binnen waren voordat dit voorstel er was: vult in
   elke regel een lege "Op de site als" en "Richting". Raakt niets anders aan. */
function vulVerhaalKolommen() {
  const blad = verhaalBlad();
  if (!blad || blad.getLastRow() < 2) { console.log('Nog geen verhalen.'); return; }
  const kop = verhaalKop(blad);
  const alumni = {};
  leesAlumni().forEach(function (a) { alumni[a.id] = a; });
  const rijen = blad.getRange(2, 1, blad.getLastRow() - 1, kop.length).getValues();
  rijen.forEach(function (rij, i) {
    const werk = String(rij[kop.indexOf('Huidige baan')] || '');
    const a = alumni[String(rij[kop.indexOf('Alumnus id')] || '').trim()] || {};
    vulLegeCellen(blad, i + 2, kop, {
      'Op de site als': rolZonderWerkgever(werk),
      'Richting': richtingVan(werk, a.interessesRuw),
    });
  });
  console.log(rijen.length + ' regel(s) nagelopen.');
}

/* Een kort bericht naar contact@, zodat geen verhaal ongezien blijft. */
function meldVerhaal(w, ingevuld, tweedeKeer, voorstel) {
  try {
    MailApp.sendEmail({
      to: ONTVANGER,
      subject: (tweedeKeer ? 'Alumniverhaal aangepast door ' : 'Alumniverhaal van ') + w['Naam'],
      name: 'Website ' + AFZENDERNAAM,
      body: [
        w['Naam'] + ' beantwoordde ' + ingevuld + ' van de ' + VERHAAL_VRAGEN.length
          + ' vragen over zijn of haar werk (' + (w['Huidige baan'] || 'baan onbekend') + ').',
        'Publiceren mag: ' + w['Mag op de site'] + '.',
        '',
        'Het script stelt voor (alleen ingevuld als de cel nog leeg was):',
        '  Op de site als: ' + ((voorstel || {})['Op de site als'] || '(niets, vul zelf in)'),
        '  Richting: ' + ((voorstel || {})['Richting'] || '(niets, vul zelf in)'),
        '',
        'Lees het in het tabblad "' + VERHAAL_TABBLAD + '". Controleer bij "Op de site als"',
        'dat er geen naam van een werkgever in staat; dat beloven we de alumni. Klopt',
        'alles, zet dan "Op de site" op yes.',
      ].join('\n'),
    });
  } catch (err) {
    console.error(err);
  }
}

/* ---- wat de site mag zien ----------------------------------------- */
function verhalenOpenbaar() {
  const blad = verhaalBlad();
  if (!blad || blad.getLastRow() < 2) return [];
  const alles = blad.getDataRange().getValues();
  const kop = alles[0].map(function (v) { return String(v).trim(); });
  const k = function (naam) { return kop.indexOf(naam); };
  const ja = function (v) { return ['yes', 'ja', 'true', 'y'].indexOf(String(v || '').trim().toLowerCase()) !== -1; };

  const uit = [];
  for (let i = 1; i < alles.length; i++) {
    const rij = alles[i];
    if (!ja(rij[k('Op de site')]) || !ja(rij[k('Mag op de site')])) continue;
    const als = String(k('Op de site als') === -1 ? '' : rij[k('Op de site als')] || '').trim();
    if (!als) continue;                 // geen anonieme omschrijving, dan niet tonen
    const antwoorden = {};
    let iets = false;
    VERHAAL_VRAGEN.forEach(function (v) {
      const t = String(k(v.kolom) === -1 ? '' : rij[k(v.kolom)] || '').trim();
      if (t) { antwoorden[v.id] = t; iets = true; }
    });
    if (!iets) continue;
    uit.push({
      als: als,
      richting: String(k('Richting') === -1 ? '' : rij[k('Richting')] || '').trim(),
      antwoorden: antwoorden,
    });
  }
  return uit;
}

function verhalenJson(vers) {
  const cache = CacheService.getScriptCache();
  if (!vers) {
    const bewaard = cache.get('verhalen');
    if (bewaard) return bewaard;
  }
  const json = JSON.stringify({ ok: true, verhalen: verhalenOpenbaar() });
  cache.put('verhalen', json, 300);
  return json;
}

/* Een wijziging in de sheet meteen op de site, zonder vijf minuten wachten. */
function wisVerhaalCache() {
  CacheService.getScriptCache().remove('verhalen');
  console.log('Cache geleegd. De site haalt de verhalen nu opnieuw op.');
}

/* ---- de mail aan de alumni ---------------------------------------- */
const VERHAAL_ONDERWERP = 'Would you share what your work is really like?';

function verhaalMailDelen(a) {
  const gegevens = [];
  if (a.werk) gegevens.push(['Current work', a.werk]);
  if (a.eerder) gegevens.push(['Previous role', a.eerder]);
  if (studieVan(a)) gegevens.push(['Studies', studieVan(a)]);
  return { voornaam: toonNaam(a.volledigeNaam) || 'there', gegevens: gegevens, link: verhaalLink(a) };
}

/* De tekst van de mail, in alinea's. Zo staat hij maar op één plek en
   kunnen de kale en de opgemaakte versie niet uit elkaar gaan lopen. */
const VERHAAL_TEKST = {
  dank: 'Thank you again for joining the Impact Connect alumni network. Because of you, '
    + 'students in Utrecht can talk to someone who has already walked the path they are curious about.',
  resultaat: 'In the last three months, we connected more than 20 students to sustainable '
    + 'opportunities! We are very happy with that, and we keep looking for what we can do better. '
    + 'One thing we noticed: students still hesitate to reach out to alumni, even when those '
    + 'alumni have very relevant experience.',
  vraag: 'That\'s why we would like to ask: would you be open to answering a few questions about '
    + 'your work, so we can publish them on our website? Answer as many or as few as you like. '
    + 'Even one answer helps, and you won\'t have to repeat anything you already told us.',
  bekend: 'In our previous form you told us:',
  bijwerken: 'Is something no longer up to date? You can change it at the top of the form.',
  knop: 'Answer the questions',
  privacy: 'We only publish your answers without your name and without the name of your '
    + 'employer, and only if you say yes in the form. You can ask us to take them down at any time.',
  motto: 'Paths are made by walking. Thank you for helping students take their first step.',
};

function verhaalMailTekst(a) {
  const d = verhaalMailDelen(a);
  const t = VERHAAL_TEKST;
  const r = ['Dear ' + d.voornaam + ',', '', t.dank, '', t.resultaat, '', t.vraag];
  if (d.gegevens.length) {
    r.push('', t.bekend);
    d.gegevens.forEach(function (g) { r.push(' * ' + g[0] + ': ' + g[1]); });
    r.push('', t.bijwerken);
  }
  r.push('', t.knop + ': ' + d.link, '', t.privacy, '', t.motto, '',
    'Warm regards,', 'The Impact Connect team');
  return r.join('\n');
}

function verhaalMailHtml(a) {
  const d = verhaalMailDelen(a);
  const t = VERHAAL_TEKST;
  const p = function (s) { return '<p>' + ontsnap(s) + '</p>'; };
  const h = ['<div style="' + MAIL_BLAD + '">'];
  h.push('<p>Dear ' + ontsnap(d.voornaam) + ',</p>');
  h.push(p(t.dank), p(t.resultaat), p(t.vraag));
  if (d.gegevens.length) {
    h.push('<p style="margin-bottom:6px">' + ontsnap(t.bekend) + '</p>');
    h.push('<ul style="padding-left:20px;margin:0 0 16px">');
    d.gegevens.forEach(function (g) {
      h.push('<li><b>' + ontsnap(g[0]) + ':</b> ' + ontsnap(g[1]) + '</li>');
    });
    h.push('</ul>');
    h.push(p(t.bijwerken));
  }
  h.push('<p style="margin:22px 0"><a href="' + ontsnap(d.link) + '" style="display:inline-block;'
    + 'background:#C2683A;color:#fff;text-decoration:none;font-weight:700;padding:12px 22px;'
    + 'border-radius:10px">' + ontsnap(t.knop) + '</a></p>');
  h.push(p(t.privacy), p(t.motto));
  h.push('<p>Warm regards,<br>The Impact Connect team</p>');
  h.push('</div>');
  return h.join('');
}

function stuurVerhaalMail(a) {
  MailApp.sendEmail({
    to: a.mail,
    subject: VERHAAL_ONDERWERP,
    name: AFZENDERNAAM,
    replyTo: ONTVANGER,
    body: verhaalMailTekst(a),
    htmlBody: verhaalMailHtml(a),
  });
}

/* Wie de mail al kreeg, per alumnus id met de datum. Staat bij het script,
   niet in de sheet, zodat niemand hem per ongeluk weghaalt. */
function verhaalVerstuurd() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty('verhaalVerstuurd') || '{}');
  } catch (e) { return {}; }
}

/* Iemand die het aanmeldformulier twee keer invulde staat er twee keer in,
   met hetzelfde mailadres en dus hetzelfde id. Die krijgt de mail één keer. */
function verhaalWachtrij() {
  const al = verhaalVerstuurd();
  const gehad = {};
  return leesAlumni().filter(function (a) {
    if (a.test || !geldigMailadres(a.mail) || al[a.id] || gehad[a.id]) return false;
    gehad[a.id] = true;
    return true;
  });
}

/* ---- de hulpfuncties voor in de editor ---------------------------- */

/* Stuurt de mail naar elk testprofiel in de alumnisheet (de naam begint met
   "Test"), zodat jullie zien hoe hij aankomt. Verandert niets. */
function testVerhaalMail() {
  const test = leesAlumni().filter(function (a) { return a.test && geldigMailadres(a.mail); });
  if (!test.length) {
    console.log('Geen testprofiel gevonden. Zet in de alumnisheet een regel waarvan de naam met "Test" begint.');
    return;
  }
  test.forEach(function (a) {
    stuurVerhaalMail(a);
    console.log('Testmail verstuurd naar ' + a.mail + '. Link: ' + verhaalLink(a));
  });
}

/* Laat zien wie de mail zou krijgen. Verstuurt niets. */
function toonVerhaalWachtrij() {
  const rij = verhaalWachtrij();
  console.log(rij.length + ' alumni zouden de mail nu krijgen:');
  rij.forEach(function (a) {
    console.log(' - ' + a.volledigeNaam + ' (' + (a.werk || 'geen baan ingevuld') + ')');
  });
  console.log('Er blijven vandaag nog ' + MailApp.getRemainingDailyQuota() + ' mails over.');
}

/* De echte ronde. Doet niets zolang VERHAAL_MAILS_AAN op false staat. Stopt
   vanzelf als het dagquotum van Google bijna op is; draai hem dan morgen
   nog een keer, dan gaat hij verder waar hij was. */
function stuurVerhaalMails() {
  if (!VERHAAL_MAILS_AAN) {
    console.log('VERHAAL_MAILS_AAN staat op false, er is niets verstuurd. Test eerst met '
      + 'testVerhaalMail(), zet hem daarna op true en push.');
    return;
  }
  const winkel = PropertiesService.getScriptProperties();
  const al = verhaalVerstuurd();
  let verstuurd = 0;
  verhaalWachtrij().some(function (a) {
    if (MailApp.getRemainingDailyQuota() < 20) {
      console.log('Dagquotum bijna op. Draai stuurVerhaalMails() morgen nog een keer.');
      return true;
    }
    try {
      stuurVerhaalMail(a);
      al[a.id] = Utilities.formatDate(new Date(), 'Europe/Amsterdam', 'yyyy-MM-dd');
      winkel.setProperty('verhaalVerstuurd', JSON.stringify(al));
      verstuurd++;
    } catch (err) {
      console.error('Mail aan ' + a.mail + ' mislukt: ' + err);
    }
    return false;
  });
  console.log(verstuurd + ' mail(s) verstuurd.');
  return verstuurd;
}

/* ---- één keer inplannen -------------------------------------------
   planVerhaalMailsMorgen() zet een wekker voor morgen om 09:30. Op dat
   moment draait verhaalMailsGepland(): die stuurt de ronde, haalt de wekker
   daarna weg en mailt contact@ hoeveel er de deur uit gingen. Bedenk je je
   nog, draai dan annuleerVerhaalMails().
   ------------------------------------------------------------------- */
const VERHAAL_PLAN_UUR = '09:30';

function planVerhaalMailsMorgen() {
  if (!VERHAAL_MAILS_AAN) {
    console.log('VERHAAL_MAILS_AAN staat op false. Er is niets ingepland.');
    return;
  }
  verhaalWekkersWeg();                      // nooit twee wekkers tegelijk
  const tz = Session.getScriptTimeZone();
  const morgen = new Date(Date.now() + 24 * 3600 * 1000);
  const dag = Utilities.formatDate(morgen, tz, 'yyyy-MM-dd');
  const verschil = Utilities.formatDate(morgen, tz, 'XXX');      // bijvoorbeeld +01:00
  const wanneer = new Date(dag + 'T' + VERHAAL_PLAN_UUR + ':00' + verschil);
  ScriptApp.newTrigger('verhaalMailsGepland').timeBased().at(wanneer).create();
  console.log('Ingepland voor ' + Utilities.formatDate(wanneer, tz, 'EEEE d MMMM yyyy, HH:mm')
    + '. ' + verhaalWachtrij().length + ' alumni krijgen dan de mail.');
}

function verhaalMailsGepland() {
  verhaalWekkersWeg();
  const verstuurd = stuurVerhaalMails() || 0;
  const over = verhaalWachtrij().length;
  try {
    MailApp.sendEmail({
      to: ONTVANGER,
      subject: 'Mail over alumniverhalen verstuurd: ' + verstuurd + ' alumni',
      name: 'Website ' + AFZENDERNAAM,
      body: [
        'De ingeplande mail over hun werk is vanochtend naar ' + verstuurd + ' alumni gegaan.',
        over ? over + ' alumni hebben hem nog niet gekregen (dagquotum of een fout). Draai '
          + 'stuurVerhaalMails() in de editor nog een keer, dan gaat hij verder.'
          : 'Iedereen op de lijst heeft hem nu.',
        '',
        'De antwoorden komen in het tabblad "' + VERHAAL_TABBLAD + '".',
      ].join('\n'),
    });
  } catch (err) {
    console.error(err);
  }
}

function annuleerVerhaalMails() {
  const weg = verhaalWekkersWeg();
  console.log(weg ? 'De ingeplande verzending is geannuleerd.' : 'Er stond niets ingepland.');
}

function verhaalWekkersWeg() {
  const oud = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'verhaalMailsGepland';
  });
  oud.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  return oud.length;
}
