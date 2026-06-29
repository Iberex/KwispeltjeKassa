/**
 * Genereert bonnen.xlsx met de juiste structuur.
 * Gebruik: node create-excel.mjs
 * Vereist: npm install exceljs  (eenmalig)
 */

import ExcelJS from 'exceljs';

// ── Pas aan naar jouw speelplein ──────────────────────────────────────────
const WEKEN = [
  'Week 1 (7–11 jul)',
  'Week 2 (14–18 jul)',
  'Week 3 (21–25 jul)',
  'Week 4 (28 jul–1 aug)',
  'Week 5 (4–8 aug)',
  'Week 6 (11–15 aug)',
];
const THEMAS = ['Jungle', 'Ruimte', 'Piraten', 'Middeleeuwen', 'Superhelden'];
// ─────────────────────────────────────────────────────────────────────────

const wb = new ExcelJS.Workbook();
wb.creator = 'Speelplein-kassa';
wb.created = new Date();

// ── Tabblad 1: Data ───────────────────────────────────────────────────────
const dataSheet = wb.addWorksheet('Data');

dataSheet.columns = [
  { header: 'Datum',        key: 'Datum',        width: 14 },
  { header: 'Week',         key: 'Week',         width: 22 },
  { header: 'Thema',        key: 'Thema',        width: 16 },
  { header: 'Winkel',       key: 'Winkel',       width: 22 },
  { header: 'Bedrag',       key: 'Bedrag',       width: 12 },
  { header: 'Omschrijving', key: 'Omschrijving', width: 36 },
  { header: 'Ingescand',    key: 'Ingescand',    width: 18 },
  { header: 'Fingerprint',  key: 'Fingerprint',  width: 34, hidden: true },
];

// Stijl headerrij
const headerRow = dataSheet.getRow(1);
headerRow.eachCell(cell => {
  cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F6EF7' } };
  cell.font   = { bold: true, color: { argb: 'FFFFFFFF' } };
  cell.border = {
    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  };
});

// Maak Excel-tabel aan (zodat de app de tabel Bonnen kan vinden)
dataSheet.addTable({
  name: 'Bonnen',
  ref: 'A1',
  headerRow: true,
  totalsRow: false,
  style: { theme: 'TableStyleMedium2', showRowStripes: true },
  columns: [
    { name: 'Datum' },
    { name: 'Week' },
    { name: 'Thema' },
    { name: 'Winkel' },
    { name: 'Bedrag' },
    { name: 'Omschrijving' },
    { name: 'Ingescand' },
    { name: 'Fingerprint' },
  ],
  rows: [],
});

// Bedrag-kolom als valuta opmaken
dataSheet.getColumn('Bedrag').numFmt = '€#,##0.00';

// ── Tabblad 2: Overzicht ──────────────────────────────────────────────────
const ovSheet = wb.addWorksheet('Overzicht');

// Titel
ovSheet.mergeCells('A1', String.fromCharCode(65 + THEMAS.length + 1) + '1');
const titleCell = ovSheet.getCell('A1');
titleCell.value = 'Overzicht per week en thema';
titleCell.font  = { size: 14, bold: true, color: { argb: 'FF4F6EF7' } };
ovSheet.getRow(1).height = 28;

// Header (rij 2): leeg | thema1 | thema2 | … | TOTAAL
const headerData = ['', ...THEMAS, 'TOTAAL'];
const headRow = ovSheet.getRow(2);
headerData.forEach((h, i) => {
  const cell = headRow.getCell(i + 1);
  cell.value = h;
  cell.font  = { bold: true, color: { argb: 'FFFFFFFF' } };
  cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F6EF7' } };
  cell.alignment = { horizontal: 'center' };
});
headRow.height = 22;
ovSheet.getColumn(1).width = 24;
THEMAS.forEach((_, i) => { ovSheet.getColumn(i + 2).width = 16; });
ovSheet.getColumn(THEMAS.length + 2).width = 14;

// Rijen per week met SUMIFS-formules
WEKEN.forEach((week, wi) => {
  const row = ovSheet.getRow(wi + 3);
  row.getCell(1).value = week;
  row.getCell(1).font  = { bold: true };

  THEMAS.forEach((thema, ti) => {
    const cell = row.getCell(ti + 2);
    // SUMIFS(Bonnen[Bedrag], Bonnen[Week], "<week>", Bonnen[Thema], "<thema>")
    cell.value = {
      formula: `SUMIFS(Bonnen[Bedrag],Bonnen[Week],"${week}",Bonnen[Thema],"${thema}")`,
    };
    cell.numFmt = '€#,##0.00';
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: wi % 2 === 0 ? 'FFF5F7FF' : 'FFFFFFFF' } };
  });

  // Rij-totaal
  const totCol = THEMAS.length + 2;
  const totCell = row.getCell(totCol);
  const firstDataCol = String.fromCharCode(66); // B
  const lastDataCol  = String.fromCharCode(65 + THEMAS.length);
  const rowNum = wi + 3;
  totCell.value = { formula: `SUM(${firstDataCol}${rowNum}:${lastDataCol}${rowNum})` };
  totCell.numFmt = '€#,##0.00';
  totCell.font   = { bold: true };
});

// Totaalrij (onderaan)
const totRow = ovSheet.getRow(WEKEN.length + 3);
totRow.getCell(1).value = 'TOTAAL';
totRow.getCell(1).font  = { bold: true, size: 12 };

THEMAS.forEach((_, ti) => {
  const col    = ti + 2;
  const letter = String.fromCharCode(64 + col);
  const cell   = totRow.getCell(col);
  cell.value   = { formula: `SUM(${letter}3:${letter}${WEKEN.length + 2})` };
  cell.numFmt  = '€#,##0.00';
  cell.font    = { bold: true };
  cell.fill    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8ECFF' } };
});

const totTotLetter = String.fromCharCode(65 + THEMAS.length + 1);
const grandCell = totRow.getCell(THEMAS.length + 2);
grandCell.value  = { formula: `SUM(B${WEKEN.length + 3}:${String.fromCharCode(65 + THEMAS.length)}${WEKEN.length + 3})` };
grandCell.numFmt = '€#,##0.00';
grandCell.font   = { bold: true, size: 12 };
grandCell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F6EF7' } };
grandCell.font   = { bold: true, color: { argb: 'FFFFFFFF' } };

totRow.height = 24;

// ── Tabblad 3: Uitleg ─────────────────────────────────────────────────────
const uitlegSheet = wb.addWorksheet('Uitleg');
uitlegSheet.getColumn(1).width = 70;

const uitlegRegels = [
  ['Speelplein-kassa — handleiding'],
  [''],
  ['HOE WERKT HET?'],
  ['1. Open de scan-app op je telefoon (link ontvangen van de organisatie).'],
  ['2. Tik "Scan een bonnetje" en neem een foto van de kassabon.'],
  ['3. Controleer de ingevulde gegevens (datum, bedrag, winkel).'],
  ['4. Kies het juiste thema en de week.'],
  ['5. Tik "Bewaar in Excel" — de bon komt automatisch hier in het tabblad Data.'],
  [''],
  ['DUBBELE BONNEN'],
  ['Als je dezelfde bon twee keer scant, krijg je een foutmelding. De bon wordt niet dubbel opgeslagen.'],
  [''],
  ['OVERZICHT'],
  ['Tabblad Overzicht toont de totalen per week en per thema. Die berekent Excel automatisch.'],
  [''],
  ['LET OP'],
  ['Verander de naam van de tabel (Bonnen) of de kolomtitels NIET.'],
  ['De app verwacht exact die namen.'],
];

uitlegRegels.forEach((r, i) => {
  const cell = uitlegSheet.getCell(`A${i + 1}`);
  cell.value = r[0];
  if (i === 0) { cell.font = { size: 14, bold: true, color: { argb: 'FF4F6EF7' } }; uitlegSheet.getRow(1).height = 26; }
  else if (r[0].match(/^[A-Z ]+$/) && r[0].length > 1) { cell.font = { bold: true }; }
});

// ── Schrijf bestand ───────────────────────────────────────────────────────
await wb.xlsx.writeFile('bonnen.xlsx');
console.log('✓ bonnen.xlsx aangemaakt. Upload dit naar OneDrive/Speelplein/');
