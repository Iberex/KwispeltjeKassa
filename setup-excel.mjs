/**
 * Bouwt bonnen.xlsx VOLLEDIG via de Microsoft Graph-API, zodat Excel Online
 * de tabel en formules native aanmaakt (gegarandeerd compatibel).
 *
 * Gebruik: node setup-excel.mjs
 * Vereist: refresh-token.txt (van get-token.mjs) in dezelfde map.
 *
 * Leest CLIENT_ID en EXCEL_PATH hieronder — pas aan indien nodig.
 */

import { readFileSync } from 'fs';
import https from 'https';

const CLIENT_ID  = 'ffc6495c-65c4-45c8-aeb3-9426a1190806';
const EXCEL_PATH = 'Zomer 26 - Kwispeltje/bonnen.xlsx';

const THEMA_NAMEN = ['Voedsel', 'Drank', 'Alcohol', 'Spelmateriaal', 'Knutselgerief', 'Kuisproducten', 'Varia'];
const WEKEN = ['Opbouw', 'Week 1', 'Week 2', 'Week 3', 'Week 4', 'Week 5', 'Week 6', 'Week 7', 'Afbraak'];
const KOLOMMEN = ['Datum', 'Week', 'Winkel', ...THEMA_NAMEN, 'Totaal', 'Fingerprint']; // 12 kolommen A..L

const refreshToken = readFileSync('refresh-token.txt', 'utf8').trim();

function req(method, host, path, headers, body) {
  return new Promise((resolve, reject) => {
    const data = body == null ? null
      : (Buffer.isBuffer(body) || typeof body === 'string' ? body : JSON.stringify(body));
    const h = { ...headers };
    if (data && !h['Content-Type']) h['Content-Type'] = 'application/json';
    if (data) h['Content-Length'] = Buffer.byteLength(data);
    const r = https.request({ method, hostname: host, path, headers: h }, res => {
      let raw = '';
      res.on('data', c => (raw += c));
      res.on('end', () => {
        const ok = res.statusCode >= 200 && res.statusCode < 300;
        let parsed; try { parsed = raw ? JSON.parse(raw) : {}; } catch { parsed = raw; }
        ok ? resolve(parsed) : reject(new Error(`${method} ${path} → ${res.statusCode}: ${raw.slice(0, 300)}`));
      });
    });
    r.on('error', reject);
    if (data) r.write(data);
    r.end();
  });
}

async function getToken() {
  const body = new URLSearchParams({
    client_id: CLIENT_ID, grant_type: 'refresh_token',
    refresh_token: refreshToken, scope: 'Files.ReadWrite offline_access',
  }).toString();
  const d = await req('POST', 'login.microsoftonline.com', '/common/oauth2/v2.0/token',
    { 'Content-Type': 'application/x-www-form-urlencoded' }, body);
  if (!d.access_token) throw new Error('Token-vernieuwing mislukt: ' + JSON.stringify(d));
  return d.access_token;
}

const colLetter = (i) => String.fromCharCode(65 + i); // 0->A

async function main() {
  const token = await getToken();
  const auth = { Authorization: 'Bearer ' + token };
  const GH = 'graph.microsoft.com';
  const padEnc = EXCEL_PATH.split('/').map(encodeURIComponent).join('/');
  const wb = `/v1.0/me/drive/root:/${padEnc}:/workbook`;

  // 1) Minimaal leeg xlsx uploaden met enkel de header-rij in "Data"
  //    (een minimaal ExcelJS/zip-bestand maken we hier via een kant-en-klare lege xlsx)
  const ExcelJS = (await import('exceljs')).default;
  const book = new ExcelJS.Workbook();
  const ws = book.addWorksheet('Data');
  ws.addRow(KOLOMMEN);
  const buf = await book.xlsx.writeBuffer();
  console.log('Uploaden basisbestand…');
  await req('PUT', GH, `/v1.0/me/drive/root:/${padEnc}:/content`,
    { ...auth, 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    Buffer.from(buf));

  // 2) Tabel "Bonnen" aanmaken over A1:L1 (Graph breidt automatisch uit bij nieuwe rijen)
  const laatsteKol = colLetter(KOLOMMEN.length - 1); // L
  console.log('Tabel Bonnen aanmaken…');
  const tabel = await req('POST', GH, `${wb}/tables/add`, auth,
    { address: `Data!A1:${laatsteKol}1`, hasHeaders: true });
  await req('PATCH', GH, `${wb}/tables/${tabel.id}`, auth, { name: 'Bonnen' });

  // Fingerprint-kolom verbergen
  await req('PATCH', GH, `${wb}/worksheets('Data')/range(address='${laatsteKol}:${laatsteKol}')`,
    auth, { columnHidden: true });

  // Euro-opmaak op bedrag-kolommen (D..K)
  await req('PATCH', GH, `${wb}/worksheets('Data')/range(address='D2:${colLetter(KOLOMMEN.length - 2)}1000')`,
    auth, { numberFormat: [['€#,##0.00']] });

  // 3) Werkblad "Overzicht" aanmaken
  console.log('Werkblad Overzicht aanmaken…');
  await req('POST', GH, `${wb}/worksheets/add`, auth, { name: 'Overzicht' });

  // Blok A2:I12 (1 header-rij + 9 weken + 1 totaalrij), 9 kolommen
  const nCols = THEMA_NAMEN.length + 2; // lege/weeklabel + categorieën + TOTAAL = 9
  const lastOvCol = colLetter(nCols - 1); // I
  const header = ['', ...THEMA_NAMEN, 'TOTAAL'];
  const rows = [header];
  WEKEN.forEach((week, wi) => {
    const r = wi + 3; // Excel-rijnummer
    const cellen = [week];
    THEMA_NAMEN.forEach((t) => cellen.push(`=SUMIFS(Bonnen[${t}],Bonnen[Week],"${week}")`));
    cellen.push(`=SUM(B${r}:${colLetter(THEMA_NAMEN.length)}${r})`); // SUM B..H
    rows.push(cellen);
  });
  // Totaalrij
  const totR = WEKEN.length + 3; // 12
  const totRij = ['TOTAAL'];
  THEMA_NAMEN.forEach((_, ti) => {
    const L = colLetter(ti + 1);
    totRij.push(`=SUM(${L}3:${L}${WEKEN.length + 2})`);
  });
  totRij.push(`=SUM(B${totR}:${colLetter(THEMA_NAMEN.length)}${totR})`);
  rows.push(totRij);

  console.log('Overzicht-formules schrijven…');
  await req('PATCH', GH, `${wb}/worksheets('Overzicht')/range(address='A2:${lastOvCol}${totR}')`,
    auth, { formulas: rows });
  // Euro-opmaak op de bedragen in Overzicht
  await req('PATCH', GH, `${wb}/worksheets('Overzicht')/range(address='B3:${lastOvCol}${totR}')`,
    auth, { numberFormat: Array(totR - 2).fill(Array(nCols - 1).fill('€#,##0.00')) });
  // Titel
  await req('PATCH', GH, `${wb}/worksheets('Overzicht')/range(address='A1')`,
    auth, { values: [['Overzicht per week en categorie']] });

  // 4) Werkblad "Uitleg"
  console.log('Werkblad Uitleg aanmaken…');
  await req('POST', GH, `${wb}/worksheets/add`, auth, { name: 'Uitleg' });
  const uitleg = [
    ['Speelplein-kassa — handleiding'],
    [''],
    ['1. Open de scan-app op je telefoon en scan een kassabon.'],
    ['2. De AI leest elk artikel en kiest automatisch een categorie.'],
    ['3. Controleer en tik "Bewaar in Excel".'],
    [''],
    ['Categorieën: ' + THEMA_NAMEN.join(', ') + '.'],
    ['Het totaal wordt over deze categorie-kolommen verdeeld.'],
    [''],
    ['Dubbele bon (winkel+datum+bedrag) wordt automatisch geweigerd.'],
    ['Tabblad Overzicht telt automatisch op per week en categorie.'],
    [''],
    ['LET OP: verander de tabelnaam (Bonnen) of de kolomtitels niet.'],
  ];
  await req('PATCH', GH, `${wb}/worksheets('Uitleg')/range(address='A1:A${uitleg.length}')`,
    auth, { values: uitleg });

  console.log('\n✓ Klaar! bonnen.xlsx is opgebouwd in OneDrive en compatibel met de app.');
}

main().catch(e => { console.error('FOUT:', e.message); process.exit(1); });
