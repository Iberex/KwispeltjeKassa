/**
 * Stap D — Eenmalig een OneDrive refresh-token ophalen.
 * Gebruik: node get-token.mjs JOUW_CLIENT_ID
 *
 * Het token wordt naar refresh-token.txt geschreven (geen copy-paste-fouten)
 * en meteen getest tegen OneDrive.
 */

import https from 'https';
import { writeFileSync } from 'fs';

const clientId = process.argv[2];
if (!clientId) {
  console.error('Gebruik: node get-token.mjs JOUW_CLIENT_ID');
  process.exit(1);
}

const scope = 'Files.ReadWrite offline_access User.Read';

function post(hostname, path, data) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams(data).toString();
    const req = https.request(
      { hostname, path, method: 'POST', headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      }},
      res => {
        let raw = '';
        res.on('data', c => (raw += c));
        res.on('end', () => resolve(JSON.parse(raw)));
      }
    );
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function graphGet(token, path) {
  return new Promise((resolve, reject) => {
    const req = https.request(
      { hostname: 'graph.microsoft.com', path, method: 'GET',
        headers: { Authorization: 'Bearer ' + token } },
      res => {
        let raw = '';
        res.on('data', c => (raw += c));
        res.on('end', () => resolve(JSON.parse(raw)));
      }
    );
    req.on('error', reject);
    req.end();
  });
}

// Stap 1: apparaatcode opvragen
const deviceRes = await post('login.microsoftonline.com', '/common/oauth2/v2.0/devicecode', {
  client_id: clientId,
  scope,
});

if (deviceRes.error) {
  console.error('Fout:', deviceRes.error_description ?? JSON.stringify(deviceRes));
  process.exit(1);
}

console.log('\n──────────────────────────────────────────────');
console.log('1. Ga naar:', deviceRes.verification_uri);
console.log('2. Voer deze code in:', deviceRes.user_code);
console.log('──────────────────────────────────────────────');
console.log('\n>>> LOG IN MET HET ACCOUNT DAT DE ONEDRIVE BEZIT (waar bonnen.xlsx staat) <<<');
console.log('\nWachten op aanmelding…\n');

// Stap 2: pollen totdat de gebruiker is aangemeld
const interval = (deviceRes.interval ?? 5) * 1000;
const expires = Date.now() + (deviceRes.expires_in ?? 900) * 1000;

let tokenRes;
while (Date.now() < expires) {
  await new Promise(r => setTimeout(r, interval));

  tokenRes = await post('login.microsoftonline.com', '/common/oauth2/v2.0/token', {
    client_id: clientId,
    grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
    device_code: deviceRes.device_code,
  });

  if (tokenRes.refresh_token) break;
  if (tokenRes.error !== 'authorization_pending') {
    console.error('Fout:', tokenRes.error_description ?? JSON.stringify(tokenRes));
    process.exit(1);
  }
}

if (!tokenRes?.refresh_token) {
  console.error('Timeout — probeer opnieuw.');
  process.exit(1);
}

console.log('✓ Aangemeld!\n');

// Stap 3: meteen testen of het token werkt tegen OneDrive
console.log('Token testen tegen OneDrive…');
const refresh = await post('login.microsoftonline.com', '/common/oauth2/v2.0/token', {
  client_id: clientId,
  grant_type: 'refresh_token',
  refresh_token: tokenRes.refresh_token,
  scope: 'Files.ReadWrite offline_access',
});

if (!refresh.access_token) {
  console.error('⚠️  Token-vernieuwing mislukt:', refresh.error_description ?? JSON.stringify(refresh));
  console.error('Het token wordt toch bewaard, maar werkt mogelijk niet.');
} else {
  const drive = await graphGet(refresh.access_token, '/v1.0/me/drive');
  if (drive.id) {
    console.log(`✓ OneDrive bereikbaar! Eigenaar: ${drive.owner?.user?.displayName ?? '?'}`);
  } else {
    console.log('⚠️  OneDrive niet bereikbaar:', JSON.stringify(drive).slice(0, 150));
  }
}

// Stap 4: token naar bestand schrijven (geen copy-paste)
writeFileSync('refresh-token.txt', tokenRes.refresh_token, 'utf8');
console.log('\n══════════════════════════════════════════════');
console.log('✓ Token bewaard in: refresh-token.txt');
console.log('══════════════════════════════════════════════');
console.log('\nLengte:', tokenRes.refresh_token.length, 'tekens');
console.log('Zet de inhoud van refresh-token.txt als MS_REFRESH_TOKEN in Vercel.\n');
