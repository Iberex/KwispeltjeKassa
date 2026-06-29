/**
 * Genereert bonnen.xlsx met de juiste structuur (categorie-kolommen).
 * Gebruik: node create-excel.mjs
 * Vereist: npm install exceljs  (eenmalig)
 */

import ExcelJS from 'exceljs';

// ── Moet exact overeenkomen met api/scan.js ────────────────────────────────
const THEMA_NAMEN = ["Voedsel", "Drank", "Alcohol", "Spelmateriaal", "Knutselgerief", "Kuisproducten", "Varia"];
const WEKEN = ["Opbouw", "Week 1", "Week 2", "Week 3", "Week 4", "Week 5", "Week 6", "Week 7", "Afbraak"];
// ───────────────────────────────────────────────────────────────────────────

// Kolommen van de tabel Bonnen (volgorde = volgorde waarin de app rijen toevoegt)
const KOLOMMEN = ["Datum", "Week", "Winkel", ...THEMA_NAMEN, "Totaal", "Fingerprint"];

const wb = new ExcelJS.Workbook();
wb.creator = 'Speelplein-kassa';
wb.created = new Date();

// ── Tabblad 1: Data ───────────────────────────────────────────────────────
const dataSheet = wb.addWorksheet('Data');

dataSheet.columns = KOLOMMEN.map((naam) => ({
  header: naam, key: naam,
  width: naam === 'Fingerprint' ? 34 : naam === 'Winkel' ? 22 : naam === 'Datum' ? 13 : 14,
  hidden: naam === 'Fingerprint',
}));

// Excel-tabel aanmaken zodat de app de tabel Bonnen kan vinden
dataSheet.addTable({
  name: 'Bonnen',
  ref: 'A1',
  headerRow: true,
  style: { theme: 'TableStyleMedium2', showRowStripes: true },
  columns: KOLOMMEN.map((naam) => ({ name: naam })),
  rows: [],
});

// Bedrag-kolommen als valuta opmaken
[...THEMA_NAMEN, 'Totaal'].forEach((naam) => {
  dataSheet.getColumn(naam).numFmt = '€#,##0.00';
});

// ── Tabblad 2: Overzicht ──────────────────────────────────────────────────
const ov = wb.addWorksheet('Overzicht');

ov.mergeCells('A1', String.fromCharCode(65 + THEMA_NAMEN.length + 1) + '1');
const titel = ov.getCell('A1');
titel.value = 'Overzicht per week en categorie';
titel.font = { size: 14, bold: true, color: { argb: 'FF2F8F5B' } };
ov.getRow(1).height = 28;

// Header (rij 2): leeg | categorie1 … | TOTAAL
const headerData = ['', ...THEMA_NAMEN, 'TOTAAL'];
const headRow = ov.getRow(2);
headerData.forEach((h, i) => {
  const cell = headRow.getCell(i + 1);
  cell.value = h;
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F8F5B' } };
  cell.alignment = { horizontal: 'center', wrapText: true };
});
headRow.height = 30;
ov.getColumn(1).width = 14;
THEMA_NAMEN.forEach((_, i) => { ov.getColumn(i + 2).width = 14; });
ov.getColumn(THEMA_NAMEN.length + 2).width = 14;

// Rijen per week met SUMIFS per categorie
WEKEN.forEach((week, wi) => {
  const rowNum = wi + 3;
  const row = ov.getRow(rowNum);
  row.getCell(1).value = week;
  row.getCell(1).font = { bold: true };

  THEMA_NAMEN.forEach((thema, ti) => {
    const cell = row.getCell(ti + 2);
    cell.value = { formula: `SUMIFS(Bonnen[${thema}],Bonnen[Week],"${week}")` };
    cell.numFmt = '€#,##0.00';
    if (wi % 2 === 0) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F2EA' } };
  });

  const firstCol = 'B';
  const lastCol = String.fromCharCode(65 + THEMA_NAMEN.length);
  const totCell = row.getCell(THEMA_NAMEN.length + 2);
  totCell.value = { formula: `SUM(${firstCol}${rowNum}:${lastCol}${rowNum})` };
  totCell.numFmt = '€#,##0.00';
  totCell.font = { bold: true };
});

// Totaalrij onderaan
const totRowNum = WEKEN.length + 3;
const totRow = ov.getRow(totRowNum);
totRow.getCell(1).value = 'TOTAAL';
totRow.getCell(1).font = { bold: true, size: 12 };

THEMA_NAMEN.forEach((_, ti) => {
  const col = ti + 2;
  const letter = String.fromCharCode(64 + col);
  const cell = totRow.getCell(col);
  cell.value = { formula: `SUM(${letter}3:${letter}${WEKEN.length + 2})` };
  cell.numFmt = '€#,##0.00';
  cell.font = { bold: true };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F2EB' } };
});

const grand = totRow.getCell(THEMA_NAMEN.length + 2);
grand.value = { formula: `SUM(B${totRowNum}:${String.fromCharCode(65 + THEMA_NAMEN.length)}${totRowNum})` };
grand.numFmt = '€#,##0.00';
grand.font = { bold: true, color: { argb: 'FFFFFFFF' } };
grand.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F8F5B' } };
totRow.height = 24;

// ── Tabblad 3: Uitleg ─────────────────────────────────────────────────────
const u = wb.addWorksheet('Uitleg');
u.getColumn(1).width = 78;

const regels = [
  ['Speelplein-kassa — handleiding'],
  [''],
  ['HOE WERKT HET?'],
  ['1. Open de scan-app op je telefoon (link van de organisatie).'],
  ['2. Tik "Scan een bonnetje" en neem een foto van de kassabon.'],
  ['3. De AI leest elk artikel uit en kiest automatisch een categorie.'],
  ['4. Controleer de categorieën (tik erop om te wijzigen) en het totaal.'],
  ['5. Tik "Bewaar in Excel" — de bon komt automatisch in het tabblad Data.'],
  [''],
  ['CATEGORIEËN'],
  ['Voedsel, Drank, Alcohol, Spelmateriaal, Knutselgerief, Kuisproducten, Varia.'],
  ['Het totaalbedrag wordt over deze categorie-kolommen verdeeld.'],
  [''],
  ['DUBBELE BONNEN'],
  ['Dezelfde bon (winkel + datum + bedrag) twee keer scannen wordt geweigerd.'],
  [''],
  ['OVERZICHT'],
  ['Tabblad Overzicht telt automatisch per week en per categorie op.'],
  [''],
  ['LET OP'],
  ['Verander de tabelnaam (Bonnen) of de kolomtitels NIET.'],
];

regels.forEach((r, i) => {
  const cell = u.getCell(`A${i + 1}`);
  cell.value = r[0];
  if (i === 0) { cell.font = { size: 14, bold: true, color: { argb: 'FF2F8F5B' } }; u.getRow(1).height = 26; }
  else if (/^[A-Z ËÉ]+$/.test(r[0]) && r[0].length > 2) { cell.font = { bold: true }; }
});

await wb.xlsx.writeFile('bonnen.xlsx');
console.log('✓ bonnen.xlsx aangemaakt met categorie-kolommen:', KOLOMMEN.join(', '));
