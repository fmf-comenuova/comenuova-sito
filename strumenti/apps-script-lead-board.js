/**
 * Come Nuova — riceve le richieste dal sito e le scrive nella tab "Leads"
 * della Sales & Operations Board.
 *
 * Installazione (5 minuti):
 * 1. Apri la board → Estensioni → Apps Script → incolla questo file.
 * 2. Distribuisci → Nuovo deployment → App web → Esegui come: te; Accesso: chiunque.
 * 3. Copia l'URL dell'app web e dammelo: lo collego al form del sito.
 *
 * Le colonne seguono l'intestazione della tab Leads (riga 6). Il sito invia solo
 * richieste già filtrate (zona e soglia minima), ma il controllo si ripete qui.
 */
const FOGLIO = 'Leads';
const RIGA_INTESTAZIONE = 6;
const ZONE_OK = ['Milano – centro e semicentro', 'Milano – resto della città', 'Provincia di Milano', 'Monza e Brianza', 'Como e provincia'];
const SOGLIA_MINIMA = 60000;

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (!ZONE_OK.includes(d.zona) || Number(d.stimaMax) < SOGLIA_MINIMA) {
    return risposta({ ok: false, motivo: 'fuori criteri' });
  }
  const sh = SpreadsheetApp.getActive().getSheetByName(FOGLIO);
  const head = sh.getRange(RIGA_INTESTAZIONE, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = n => head.indexOf(n);
  const riga = new Array(head.length).fill('');
  const ultimo = sh.getRange(sh.getLastRow(), col('Lead_ID') + 1).getValue();
  const num = (parseInt(String(ultimo).replace(/\D/g, ''), 10) || 0) + 1;

  const set = (n, v) => { if (col(n) >= 0) riga[col(n)] = v; };
  set('Lead_ID', 'LEAD_' + String(num).padStart(4, '0'));
  set('Client', d.nome);
  set('Source', 'Direct');
  set('Source Name', 'Sito – stimatore');
  set('Address', d.indirizzo);
  set('City', d.zona.startsWith('Como') ? 'CO' : d.zona.startsWith('Monza') ? 'MB' : 'MI');
  set('Sqm', d.mq);
  set('Type of Reno', 'Come Nuova');
  set('Budget', `€ ${Math.round(d.stimaMin / 1000)}–${Math.round(d.stimaMax / 1000)}k (${d.livello})`);
  set('Expected SoW', d.descrizione);
  set('Lead Status', 'WIP');
  set('Architect', d.architetto === 'Sì' ? 'Ha già architetto' : d.architetto);
  sh.appendRow(riga);

  MailApp.sendEmail({
    to: 'fmf@comenuova.com',
    subject: `Nuova lead dal sito: ${d.mq} mq, ${d.zona}, ${d.livello}`,
    body: `${d.nome} · ${d.email} · ${d.telefono}\n${d.indirizzo}\nInizio: ${d.quando}\n\n${d.descrizione}`
  });
  return risposta({ ok: true });
}

function risposta(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
