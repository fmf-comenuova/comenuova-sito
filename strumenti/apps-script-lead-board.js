/**
 * Come Nuova — riceve le richieste dal sito e le scrive nella tab "Leads"
 * della Sales & Operations Board, poi avvisa Francesco e risponde al cliente.
 *
 * Installazione (5 minuti):
 * 1. Apri la board → Estensioni → Apps Script → incolla questo file (sostituisci tutto).
 * 2. Nella tab Leads aggiungi due colonne in fondo all'intestazione (riga 6): "Email" e "Phone"
 *    (a destra dell'ultima colonna). Senza, email e telefono arrivano solo per mail.
 * 3. Distribuisci → Nuovo deployment → App web → Esegui come: te; Accesso: chiunque.
 *    Autorizza i permessi quando richiesto (fogli, invio email).
 * 4. Copia l'URL dell'app web (finisce con /exec) e mandalo a Claude: lo collega al form del sito.
 *
 * Dopo ogni modifica a questo file: Distribuisci → Gestisci deployment → modifica → Nuova versione.
 * Le colonne seguono l'intestazione della tab Leads (riga 6). Il sito invia solo richieste già
 * filtrate (zona e soglia minima), ma il controllo si ripete qui.
 */
const FOGLIO = 'Leads';
const RIGA_INTESTAZIONE = 6;
const ZONE_OK = ['Milano – centro e semicentro', 'Milano – resto della città', 'Provincia di Milano', 'Monza e Brianza', 'Como e provincia'];
const SOGLIA_MINIMA = 60000;
const EMAIL_FRANCESCO = 'fmf@comenuova.com';

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return risposta({ ok: false, motivo: 'dati non validi' }); }

  const emailOk = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(d.email || ''));
  if (!d.consenso || !emailOk || String(d.descrizione || '').trim().length < 60 || !String(d.nome || '').trim()) {
    return risposta({ ok: false, motivo: 'dati incompleti' });
  }
  if (!ZONE_OK.includes(d.zona) || Number(d.stimaMax) < SOGLIA_MINIMA) {
    return risposta({ ok: false, motivo: 'fuori criteri' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let leadId;
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(FOGLIO);
    const head = sh.getRange(RIGA_INTESTAZIONE, 1, 1, sh.getLastColumn()).getValues()[0];
    const col = n => head.indexOf(n);

    // Prossimo Lead_ID: massimo numero presente nella colonna, non l'ultima riga.
    const ids = sh.getRange(RIGA_INTESTAZIONE + 1, col('Lead_ID') + 1, Math.max(sh.getLastRow() - RIGA_INTESTAZIONE, 1), 1).getValues();
    const max = ids.reduce((m, r) => Math.max(m, parseInt(String(r[0]).replace(/\D/g, ''), 10) || 0), 0);
    leadId = 'LEAD_' + String(max + 1).padStart(4, '0');

    const riga = new Array(head.length).fill('');
    const set = (n, v) => { if (col(n) >= 0) riga[col(n)] = sicuro(v); };
    const fonte = d.utm_source ? `Sito – ${d.utm_source}${d.utm_campaign ? ' / ' + d.utm_campaign : ''}` : 'Sito – stimatore';

    set('Lead_ID', leadId);
    set('Client', d.nome);
    set('Source', 'Direct');
    set('Source Name', fonte);
    set('Address', d.indirizzo);
    set('City', d.zona.startsWith('Como') ? 'CO' : d.zona.startsWith('Monza') ? 'MB' : 'MI');
    set('Sqm', d.mq);
    set('Type of Reno', 'Come Nuova');
    set('Budget', `€ ${Math.round(d.stimaMin / 1000)}–${Math.round(d.stimaMax / 1000)}k (${d.livello})`);
    set('Expected SoW', d.descrizione);
    set('Lead Status', 'WIP');
    set('Architect', d.architetto === 'Sì' ? 'Ha già architetto' : d.architetto);
    set('Email', d.email);
    set('Phone', d.telefono);
    sh.appendRow(riga);
  } finally {
    lock.releaseLock();
  }

  const forchetta = `€ ${Math.round(d.stimaMin / 1000)}–${Math.round(d.stimaMax / 1000)}k`;

  MailApp.sendEmail({
    to: EMAIL_FRANCESCO,
    replyTo: d.email,
    subject: `Nuova lead dal sito (${leadId}): ${d.mq} mq, ${d.zona}, ${d.livello}`,
    body: `${d.nome} · ${d.email} · ${d.telefono}\n${d.indirizzo}\nInizio: ${d.quando}\nStima mostrata: ${forchetta}\nFonte: ${d.utm_source || 'sito'} ${d.utm_campaign || ''}\n\n${d.descrizione}`
  });

  MailApp.sendEmail({
    to: d.email,
    replyTo: EMAIL_FRANCESCO,
    name: 'Come Nuova',
    subject: `La tua ristrutturazione in ${d.indirizzo} — prossimi passi`,
    body:
`Buongiorno ${String(d.nome).split(' ')[0]},

grazie per averci scritto. Abbiamo ricevuto la descrizione del progetto in ${d.indirizzo}:

"${d.descrizione}"

La stima indicativa che hai visto sul sito è ${forchetta}. Il prezzo vero arriva dal computo metrico, dopo il sopralluogo.

Ti scriviamo entro un giorno lavorativo per proporti una call di 20 minuti. In quell'occasione ti chiederemo:
- planimetria attuale, se ce l'hai;
- eventuale progetto dell'architetto;
- quando prevedi di iniziare.

Se nel frattempo vuoi vedere come lavoriamo, qui trovi alcuni progetti:
https://www.comenuova.com/#progetti

A presto,

Francesco Maria Fama - +39 339 6290314`
  });

  return risposta({ ok: true });
}

// Evita che testo che inizia con = + - @ venga interpretato come formula dal foglio.
function sicuro(v) {
  v = (v === undefined || v === null) ? '' : v;
  if (typeof v !== 'string') return v;
  v = v.slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function risposta(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
