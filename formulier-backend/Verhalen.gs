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
     stuurVerhaalMails()     echt naar iedereen. Werkt pas als
                             VERHAAL_MAILS_AAN hieronder op true staat.
   Wie de mail al kreeg, krijgt hem nooit een tweede keer.
   =================================================================== */

/* Staat bewust uit. Pas op true zetten (en pushen) als jullie de mail
   hebben getest en akkoord zijn. Dan pas doet stuurVerhaalMails() iets. */
const VERHAAL_MAILS_AAN = false;

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
    } else {
      const kop = verhaalKop(blad);
      blad.appendRow(kop.map(function (naam) {
        return waarden[naam] === undefined ? '' : waarden[naam];
      }));
    }
    gelukt = true;
  });
  if (!gelukt) return antwoord({ ok: false, reden: 'sheet' });

  meldVerhaal(waarden, ingevuld, tweedeKeer);
  return antwoord({ ok: true });
}

/* Een kort bericht naar contact@, zodat geen verhaal ongezien blijft. */
function meldVerhaal(w, ingevuld, tweedeKeer) {
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
        'Lees het in het tabblad "' + VERHAAL_TABBLAD + '". Wil je het op de site, zet dan',
        '"Op de site" op yes en typ bij "Op de site als" een anonieme omschrijving,',
        'zonder naam of werkgever, bijvoorbeeld "Data analyst at a large consultancy".',
        'Bij "Richting" kun je het een groep geven, zoals Consultancy of Data.',
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

function verhaalMailTekst(a) {
  const d = verhaalMailDelen(a);
  const r = [
    'Dear ' + d.voornaam + ',',
    '',
    'Thank you again for joining the Impact Connect alumni network. Because of you,',
    'students in Utrecht can talk to someone who has already walked the path they',
    'are curious about.',
    '',
    'Many students have no idea what a job like yours actually looks like: what a',
    'week is like, how much you work with people, and how you got there. We would',
    'love to share that on our website, anonymously, so more students get a feel',
    'for it.',
  ];
  if (d.gegevens.length) {
    r.push('', 'In our previous form you told us:');
    d.gegevens.forEach(function (g) { r.push(' * ' + g[0] + ': ' + g[1]); });
  }
  r.push('',
    'Would you be open to answering a few questions about this work? Answer as many',
    'or as few as you like. Even one answer helps, and you won\'t have to repeat',
    'anything you already told us.',
    '',
    'Answer the questions: ' + d.link,
    '',
    'Is something above out of date? You can correct it at the top of the form.',
    '',
    'We only publish your answers without your name and without the name of your',
    'employer, and only if you say yes in the form. You can ask us to take them',
    'down at any time.',
    '',
    'Paths are made by walking. Thank you for helping students take their first step.',
    '',
    'Warm regards,',
    'The Impact Connect team');
  return r.join('\n');
}

function verhaalMailHtml(a) {
  const d = verhaalMailDelen(a);
  const h = ['<div style="' + MAIL_BLAD + '">'];
  h.push('<p>Dear ' + ontsnap(d.voornaam) + ',</p>');
  h.push('<p>Thank you again for joining the Impact Connect alumni network. Because of you, '
    + 'students in Utrecht can talk to someone who has already walked the path they are curious about.</p>');
  h.push('<p>Many students have no idea what a job like yours actually looks like: what a week is '
    + 'like, how much you work with people, and how you got there. We would love to share that on '
    + 'our website, anonymously, so more students get a feel for it.</p>');
  if (d.gegevens.length) {
    h.push('<p style="margin-bottom:6px">In our previous form you told us:</p>');
    h.push('<ul style="padding-left:20px;margin:0 0 16px">');
    d.gegevens.forEach(function (g) {
      h.push('<li><b>' + ontsnap(g[0]) + ':</b> ' + ontsnap(g[1]) + '</li>');
    });
    h.push('</ul>');
  }
  h.push('<p>Would you be open to answering a few questions about this work? Answer as many or as '
    + 'few as you like. Even one answer helps, and you won&#39;t have to repeat anything you already told us.</p>');
  h.push('<p style="margin:22px 0"><a href="' + ontsnap(d.link) + '" style="display:inline-block;'
    + 'background:#C2683A;color:#fff;text-decoration:none;font-weight:700;padding:12px 22px;'
    + 'border-radius:10px">Answer the questions</a></p>');
  h.push('<p>Is something above out of date? You can correct it at the top of the form.</p>');
  h.push('<p>We only publish your answers without your name and without the name of your employer, '
    + 'and only if you say yes in the form. You can ask us to take them down at any time.</p>');
  h.push('<p>Paths are made by walking. Thank you for helping students take their first step.</p>');
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
}
