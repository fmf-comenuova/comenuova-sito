# comenuova.com

Sito di Come Nuova S.r.l. Pubblicato su Netlify: ogni modifica al ramo `main` va online da sola.

- `index.html`: sito (anteprima, con `noindex` finché non si lancia)
- `img/`: foto e planimetrie dei progetti
- `fonts/`: Neue Haas Grotesk (licenza web MyFonts intestata a Come Nuova, solo per comenuova.com)
- `strumenti/`: script Google Apps Script per salvare le lead nella board, testi delle risposte automatiche

## Da compilare al lancio

In `index.html`, oggetto `CONFIG` (in cima allo script):

- `SCRIPT_URL`: URL dell'app web Apps Script (vedi `strumenti/apps-script-lead-board.js`).
- `CAL_URL`: link alla pagina di prenotazione di Google Calendar.
- `PIXEL_ID`: ID del Meta Pixel. Se vuoto, banner cookie e pixel restano spenti.

Al lancio: togliere `noindex` e il banner "Bozza", far verificare l'informativa `#privacy`, puntare il dominio su Netlify.
Link per gli ads: `https://www.comenuova.com/?utm_source=instagram&utm_medium=paid&utm_campaign=NOME&utm_content=CREATIVITA#stima`
