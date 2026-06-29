// ── Aanpassen aan jouw speelplein ──────────────────────────────────────────
export const WEEK_RANGES = [
  { naam: 'Week 1 (7–11 jul)',     van: '2025-07-07', tot: '2025-07-11' },
  { naam: 'Week 2 (14–18 jul)',    van: '2025-07-14', tot: '2025-07-18' },
  { naam: 'Week 3 (21–25 jul)',    van: '2025-07-21', tot: '2025-07-25' },
  { naam: 'Week 4 (28 jul–1 aug)', van: '2025-07-28', tot: '2025-08-01' },
  { naam: 'Week 5 (4–8 aug)',      van: '2025-08-04', tot: '2025-08-08' },
  { naam: 'Week 6 (11–15 aug)',    van: '2025-08-11', tot: '2025-08-15' },
];

export const THEMA_NAMEN = [
  'Jungle',
  'Ruimte',
  'Piraten',
  'Middeleeuwen',
  'Superhelden',
];
// ───────────────────────────────────────────────────────────────────────────

export const config = { api: { bodyParser: { sizeLimit: '12mb' } } };

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { imageBase64, mimeType = 'image/jpeg' } = req.body ?? {};
  if (!imageBase64) return res.status(400).json({ error: 'Geen afbeelding meegestuurd.' });

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'GEMINI_API_KEY niet ingesteld.' });

  // AQ.-sleutels zijn auth keys en vereisen Bearer-authenticatie
  const geminiHeaders = { 'Content-Type': 'application/json' };
  const geminiUrl = apiKey.startsWith('AQ.')
    ? 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent'
    : `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  if (apiKey.startsWith('AQ.')) geminiHeaders['Authorization'] = `Bearer ${apiKey}`;

  const geminiRes = await fetch(geminiUrl, {
      method: 'POST',
      headers: geminiHeaders,
      body: JSON.stringify({
        contents: [{
          parts: [
            {
              text: `Je bent een kassabon-scanner. Analyseer deze kassabon zorgvuldig.
Geef de volgende informatie terug als pure JSON (geen markdown, geen uitleg):
{
  "datum": "YYYY-MM-DD",
  "winkel": "naam van de winkel",
  "bedrag": 12.34,
  "omschrijving": "korte omschrijving van de aankopen (max 120 tekens)"
}
Als datum niet leesbaar is, gebruik vandaag: ${new Date().toISOString().slice(0, 10)}.
Als bedrag niet leesbaar is, gebruik null.
Gebruik altijd het TOTAALBEDRAG (inclusief btw).`,
            },
            {
              inline_data: { mime_type: mimeType, data: imageBase64 },
            },
          ],
        }],
        generationConfig: { temperature: 0, maxOutputTokens: 300 },
      }),
    },
  );

  if (!geminiRes.ok) {
    const err = await geminiRes.text();
    return res.status(502).json({ error: 'Gemini-fout: ' + err.slice(0, 200) });
  }

  const geminiData = await geminiRes.json();
  const raw = geminiData.candidates?.[0]?.content?.parts?.[0]?.text ?? '';

  let extracted;
  try {
    const jsonMatch = raw.match(/\{[\s\S]*\}/);
    extracted = JSON.parse(jsonMatch?.[0] ?? raw);
  } catch {
    return res.status(422).json({ error: 'Kon de bon niet uitlezen. Probeer een scherpere foto.' });
  }

  // Bepaal de week op basis van de datum
  const bonDatum = extracted.datum ? new Date(extracted.datum) : new Date();
  const week = WEEK_RANGES.find(w => {
    const van = new Date(w.van);
    const tot = new Date(w.tot);
    tot.setHours(23, 59, 59);
    return bonDatum >= van && bonDatum <= tot;
  })?.naam ?? 'Buiten schema';

  return res.status(200).json({
    datum: extracted.datum ?? new Date().toISOString().slice(0, 10),
    winkel: extracted.winkel ?? '',
    bedrag: extracted.bedrag ?? null,
    omschrijving: extracted.omschrijving ?? '',
    week,
    themas: THEMA_NAMEN,
    weekRanges: WEEK_RANGES.map(w => w.naam),
  });
}
