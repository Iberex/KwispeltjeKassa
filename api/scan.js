// Serverless functie (Vercel). Twee acties:
//  - action "extract": leest de bon met Groq (gratis AI) en checkt of hij al bestaat
//  - action "save": voegt de (gecontroleerde) bon toe als rij in de Excel-tabel
//
// Vereiste omgevingsvariabelen (env vars), in te stellen in Vercel:
//  GROQ_API_KEY      - gratis sleutel van console.groq.com
//  MS_CLIENT_ID      - Application (client) ID uit de Azure app-registratie
//  MS_REFRESH_TOKEN  - eenmalig opgehaald met get-token.mjs
//  EXCEL_PATH        - pad naar het bestand in OneDrive, bv. "Zomer 26 - Kwispeltje/bonnen.xlsx"
//  TABLE_NAME        - naam van de tabel in Excel, bv. "Bonnen"

export const config = { api: { bodyParser: { sizeLimit: '12mb' } } };

const THEMA_NAMEN = ["Voedsel", "Drank", "Alcohol", "Spelmateriaal", "Knutselgerief", "Kuisproducten", "Varia"];

const WEEK_RANGES = [
  { naam: "Opbouw", start: "2026-06-29", eind: "2026-07-05" },
  { naam: "Week 1", start: "2026-07-06", eind: "2026-07-12" },
  { naam: "Week 2", start: "2026-07-13", eind: "2026-07-19" },
  { naam: "Week 3", start: "2026-07-20", eind: "2026-07-26" },
  { naam: "Week 4", start: "2026-07-27", eind: "2026-08-02" },
  { naam: "Week 5", start: "2026-08-03", eind: "2026-08-09" },
  { naam: "Week 6", start: "2026-08-10", eind: "2026-08-16" },
  { naam: "Week 7", start: "2026-08-17", eind: "2026-08-23" },
  { naam: "Afbraak", start: "2026-08-24", eind: "2026-08-30" },
];
function weekVanDatum(d) {
  if (!d) return "Week 1";
  const eerste = WEEK_RANGES[0], laatste = WEEK_RANGES[WEEK_RANGES.length - 1];
  if (d < eerste.start) return eerste.naam;
  if (d > laatste.eind) return laatste.naam;
  const r = WEEK_RANGES.find((x) => d >= x.start && d <= x.eind);
  return r ? r.naam : "Week 1";
}

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

function berekenVerdeling(items, totaal) {
  const sums = {}; let itemSom = 0;
  (items || []).forEach((it) => {
    const t = THEMA_NAMEN.includes(it.thema) ? it.thema : "Varia";
    const b = Number(it.bedrag) || 0; sums[t] = (sums[t] || 0) + b; itemSom += b;
  });
  const v = {}; THEMA_NAMEN.forEach((t) => (v[t] = 0));
  if (itemSom > 0) THEMA_NAMEN.forEach((t) => { if (sums[t]) v[t] = r2(sums[t] / itemSom * (Number(totaal) || 0)); });
  else v["Varia"] = r2(totaal);
  return v;
}

const fingerprint = (winkel, datum, totaal) =>
  `${(winkel || "").trim().toLowerCase()}|${datum}|${r2(totaal)}`;

// ---- Groq (gratis AI met beeldherkenning) ----
const PROMPT = `Je leest kassabonnen van een Vlaams speelplein uit. Geef ENKEL geldige JSON terug, zonder uitleg en zonder markdown.
{
 "winkel": string,
 "datum": "YYYY-MM-DD" of null (datum op de bon),
 "items": [{"omschrijving": string, "bedrag": number, "thema": een thema}],
 "totaal": number (eindtotaal in euro),
 "leesbaar": true of false,
 "opmerking": korte tekst, leeg als alles duidelijk is
}
Bepaal per item zelf het thema op basis van wat het product echt is. Mogelijke thema's:
Voedsel (eten: brood, skyr, groenten, snoep, vlees, kaas, zuivel...), Drank (frisdrank, water, sap, koffie - geen alcohol),
Alcohol (bier, wijn, sterke drank), Spelmateriaal (speelgoed, sport, ballen),
Knutselgerief (verf, papier, lijm, stiften), Kuisproducten (Dreft, allesreiniger, vuilniszakken, handzeep),
Varia (de rest). Voorbeelden: "Dreft 1L"->Kuisproducten, "Skyr 390g"->Voedsel, "Cola"->Drank, "Jupiler"->Alcohol.
Lees max 20 lijnen. Punt als decimaalteken. Wazig/onleesbaar? leesbaar=false en leg uit in opmerking.`;

async function aiLees(base64, mimeType) {
  const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + process.env.GROQ_API_KEY,
    },
    body: JSON.stringify({
      model: "meta-llama/llama-4-scout-17b-16e-instruct",
      temperature: 0,
      max_tokens: 1200,
      response_format: { type: "json_object" },
      messages: [{
        role: "user",
        content: [
          { type: "text", text: PROMPT },
          { type: "image_url", image_url: { url: `data:${mimeType};base64,${base64}` } },
        ],
      }],
    }),
  });
  if (!res.ok) throw new Error("AI-uitlezing mislukt (" + res.status + ").");
  const data = await res.json();
  const tekst = data?.choices?.[0]?.message?.content || "";
  return JSON.parse(tekst.replace(/```json|```/g, "").trim());
}

// ---- Microsoft Graph ----
async function graphToken() {
  const res = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.MS_CLIENT_ID, grant_type: "refresh_token",
      refresh_token: process.env.MS_REFRESH_TOKEN, scope: "Files.ReadWrite offline_access",
    }),
  });
  const d = await res.json();
  if (!d.access_token) throw new Error("OneDrive-token vernieuwen mislukt: " + (d.error_description || JSON.stringify(d)));
  return d.access_token;
}
function tabelBasis() {
  const pad = process.env.EXCEL_PATH.split("/").map(encodeURIComponent).join("/");
  return `https://graph.microsoft.com/v1.0/me/drive/root:/${pad}:/workbook/tables/${encodeURIComponent(process.env.TABLE_NAME)}`;
}
async function bestaandeVingerafdrukken(token) {
  const res = await fetch(tabelBasis() + "/rows", { headers: { Authorization: "Bearer " + token } });
  if (!res.ok) throw new Error("Excel-tabel lezen mislukt (" + res.status + "). Klopt EXCEL_PATH en TABLE_NAME?");
  const d = await res.json();
  // laatste kolom = Fingerprint
  return new Set((d.value || []).map((r) => String(r.values?.[0]?.slice(-1)?.[0] ?? "")));
}
async function voegRijToe(token, rij) {
  const res = await fetch(tabelBasis() + "/rows/add", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ values: [rij] }),
  });
  if (!res.ok) throw new Error("Rij toevoegen aan Excel mislukt (" + res.status + ").");
}

function bonUitInvoer(body) {
  const datum = body.datum || new Date().toISOString().slice(0, 10);
  const items = Array.isArray(body.items) ? body.items : [];
  const totaal = r2(body.totaal);
  return {
    winkel: body.winkel || "Onbekende winkel", datum,
    week: body.week || weekVanDatum(datum), totaal,
    verdeling: berekenVerdeling(items, totaal),
    fp: fingerprint(body.winkel, datum, totaal),
  };
}

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ fout: "Alleen POST." });
  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    const { action } = body;

    if (action === "extract") {
      if (!body.image) return res.status(400).json({ fout: "Geen foto ontvangen." });
      const r = await aiLees(body.image, body.mimeType || "image/jpeg");
      const datum = r.datum || new Date().toISOString().slice(0, 10);
      const items = (Array.isArray(r.items) ? r.items : []).map((it) => ({
        omschrijving: it.omschrijving || "Artikel", bedrag: Number(it.bedrag) || 0,
        thema: THEMA_NAMEN.includes(it.thema) ? it.thema : "Varia",
      }));
      const totaal = Number(r.totaal) || items.reduce((s, it) => s + it.bedrag, 0);
      const fp = fingerprint(r.winkel, datum, totaal);
      let dubbel = false;
      try { dubbel = (await bestaandeVingerafdrukken(await graphToken())).has(fp); } catch (e) { /* dedup niet kritiek bij extract */ }
      return res.status(200).json({
        winkel: r.winkel || "Onbekende winkel", datum, items, totaal,
        week: weekVanDatum(datum), leesbaar: r.leesbaar !== false, opmerking: r.opmerking || "",
        dubbel,
      });
    }

    if (action === "save") {
      const bon = bonUitInvoer(body);
      const token = await graphToken();
      const bestaand = await bestaandeVingerafdrukken(token);
      if (bestaand.has(bon.fp)) return res.status(200).json({ status: "dubbel" });
      const rij = [
        bon.datum, bon.week, bon.winkel,
        ...THEMA_NAMEN.map((t) => bon.verdeling[t] || 0),
        bon.totaal, bon.fp,
      ];
      await voegRijToe(token, rij);
      return res.status(200).json({ status: "ok" });
    }

    return res.status(400).json({ fout: "Onbekende actie." });
  } catch (e) {
    return res.status(500).json({ fout: e.message || "Er ging iets mis." });
  }
}
