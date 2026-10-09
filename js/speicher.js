// Speichert den Stand auf diesem Gerät (im Browser) und – am PC – als Excel im Ordner „Daten“.
const KEY = "kisy-v1", KEY_VORHER = "kisy-v1-vorher", KEY_BASIS = "kisy-v1-basis";

export function laden() {
  try { const s = localStorage.getItem(KEY); return s ? JSON.parse(s) : null; } catch (e) { return null; }
}
export function speichern(daten) {
  try { localStorage.setItem(KEY, JSON.stringify(daten)); return true; } catch (e) { return false; }
}
// Vor dem Einlesen einer Excel-Datei den jetzigen Stand aufheben (für „Vorherigen Stand wiederherstellen“).
export function sichern() {
  try { const s = localStorage.getItem(KEY); if (s) localStorage.setItem(KEY_VORHER, s); } catch (e) { /* egal */ }
}
export function vorherLaden() {
  try { const s = localStorage.getItem(KEY_VORHER); return s ? JSON.parse(s) : null; } catch (e) { return null; }
}
export function allesLoeschen() {
  try { for (const k of [KEY, KEY_VORHER, KEY_BASIS]) localStorage.removeItem(k); } catch (e) { /* egal */ }
}
// Stand beim letzten Abgleich mit Google Drive (Grundlage für das Zusammenführen, siehe abgleich.js)
export function basisLaden() {
  try { const s = localStorage.getItem(KEY_BASIS); return s ? JSON.parse(s) : null; } catch (e) { return null; }
}
export function basisSpeichern(basis) {
  try { localStorage.setItem(KEY_BASIS, JSON.stringify(basis)); return true; } catch (e) { return false; }
}
// Den Browser bitten, die Daten nicht von selbst zu löschen (z. B. bei wenig Speicherplatz).
export async function dauerhaftBitten() {
  try { if (navigator.storage && navigator.storage.persist) return await navigator.storage.persist(); } catch (e) { /* egal */ }
  return false;
}

/* ---------- Ordner auf dem PC (nur wenn die App über werkzeuge/server.js läuft) ---------- */
const H = { "X-KISY": "1" };
export async function ordnerVerfuegbar() {
  if (!/^(localhost|127\.0\.0\.1)$/.test(location.hostname)) return null;
  try { const r = await fetch("lokal/info", { headers: H }); return r.ok ? await r.json() : null; } catch (e) { return null; }
}
export async function ausOrdnerLaden() {
  const r = await fetch("lokal/stand", { headers: H });
  if (!r.ok) throw new Error(await r.text());
  return { name: decodeURIComponent(r.headers.get("X-Dateiname") || ""), daten: await r.arrayBuffer() };
}
export async function inOrdnerSpeichern(daten) {
  const r = await fetch("lokal/stand", { method: "POST", headers: H, body: daten });
  if (!r.ok) throw new Error(await r.text());
  return (await r.json()).datei;
}
