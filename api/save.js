import { createHash } from 'crypto';

export const config = { api: { bodyParser: { sizeLimit: '1mb' } } };

const MS_TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
const GRAPH = 'https://graph.microsoft.com/v1.0';

async function getAccessToken() {
  const res = await fetch(MS_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: process.env.MS_CLIENT_ID,
      grant_type: 'refresh_token',
      refresh_token: process.env.MS_REFRESH_TOKEN,
      scope: 'Files.ReadWrite offline_access',
    }),
  });
  const data = await res.json();
  if (!data.access_token) throw new Error('Token-vernieuwing mislukt: ' + JSON.stringify(data));
  return data.access_token;
}

async function graphGet(token, path) {
  const res = await fetch(GRAPH + path, {
    headers: { Authorization: 'Bearer ' + token },
  });
  if (!res.ok) throw new Error(`Graph GET ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphPost(token, path, body) {
  const res = await fetch(GRAPH + path, {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Graph POST ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { datum, week, thema, winkel, bedrag, omschrijving } = req.body ?? {};

  if (!datum || bedrag === undefined || bedrag === null) {
    return res.status(400).json({ error: 'Datum en bedrag zijn verplicht.' });
  }

  const fingerprint = createHash('sha256')
    .update(`${datum}|${(winkel ?? '').toLowerCase().trim()}|${bedrag}`)
    .digest('hex')
    .slice(0, 32);

  const ingescand = new Date().toISOString().slice(0, 19).replace('T', ' ');
  const excelPath = process.env.EXCEL_PATH ?? 'Speelplein/bonnen.xlsx';
  const tableName = process.env.TABLE_NAME ?? 'Bonnen';

  let token;
  try {
    token = await getAccessToken();
  } catch (e) {
    return res.status(502).json({ error: 'OneDrive-verbinding mislukt: ' + e.message });
  }

  // Haal bestand-ID op
  let fileId;
  try {
    const fileInfo = await graphGet(token, `/me/drive/root:/${excelPath}`);
    fileId = fileInfo.id;
  } catch (e) {
    return res.status(404).json({ error: `Bestand '${excelPath}' niet gevonden in OneDrive.` });
  }

  const tableBase = `/me/drive/items/${fileId}/workbook/tables/${tableName}`;

  // Controleer op dubbele bon via de Fingerprint-kolom
  let rows;
  try {
    const rowsData = await graphGet(token, `${tableBase}/rows?$select=values`);
    rows = rowsData.value ?? [];
  } catch (e) {
    return res.status(502).json({ error: 'Kon tabelrijen niet ophalen: ' + e.message });
  }

  // Kolom 7 (index 7) is Fingerprint (0-gebaseerd: Datum=0 … Fingerprint=7)
  const isDuplicate = rows.some(r => r.values?.[0]?.[7] === fingerprint);
  if (isDuplicate) {
    return res.status(409).json({ error: 'Deze bon is al ingescand (dubbel geweigerd).' });
  }

  // Voeg rij toe
  // Volgorde moet overeenkomen met de tabelkolommen in bonnen.xlsx:
  // Datum | Week | Thema | Winkel | Bedrag | Omschrijving | Ingescand | Fingerprint
  try {
    await graphPost(token, `${tableBase}/rows/add`, {
      values: [[datum, week ?? '', thema ?? '', winkel ?? '', bedrag, omschrijving ?? '', ingescand, fingerprint]],
    });
  } catch (e) {
    return res.status(502).json({ error: 'Opslaan in Excel mislukt: ' + e.message });
  }

  return res.status(200).json({ ok: true });
}
