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

  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'GROQ_API_KEY niet ingesteld.' });

  const prompt = `Je bent een kassabon-scanner. Analyseer deze kassabon zorgvuldig.
Geef UITSLUITEND pure JSON terug (geen markdown, geen uitleg):
{
  "datum": "YYYY-MM-DD",
  "winkel": "naam van de winkel",
  "bedrag": 12.34,
  "omschrijving": "korte omschrijving van de aankopen (max 120 tekens)"
}
Als datum niet leesbaar is, gebruik vandaag: ${new Date().toISOString().slice(0, 10)}.
Als bedrag niet leesbaar is, gebruik null.
Gebruik altijd het TOTAALBEDRAG (inclusief btw).`;

  const geminiRes = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: 'meta-llama/llama-4-scout-17b-16e-instruct',
      temperature: 0,
      max_tokens: 400,
      messages: [{
        role: 'user',
        content: [
          { type: 'text', text: prompt },
          { type: 'image_url', image_url: { url: `data:${mimeType};base64,${imageBase64}` } },
        ],
      }],
    }),
  });

  if (!geminiRes.ok) {
    const err = await geminiRes.text();
    return res.status(502).json({ error: 'AI-fout: ' + err.slice(0, 200) });
  }

  const geminiData = await geminiRes.json();
  const raw = geminiData.choices?.[0]?.message?.content ?? '';

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
