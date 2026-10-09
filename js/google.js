// Anmeldung bei Google und Zugriff auf Google Drive – direkt vom Gerät aus, ohne eigenen Server.
// Recht: nur Dateien, die KISY selbst angelegt hat (drive.file). Der Zugangsschlüssel bleibt auf diesem Gerät.
import { GOOGLE_CLIENT_ID } from "./konfig.js";

const SCOPES = "https://www.googleapis.com/auth/drive.file";
const KEY = "kisy-google";
const DRIVE = "https://www.googleapis.com/drive/v3/files";
const UPLOAD = "https://www.googleapis.com/upload/drive/v3/files";
export const ORDNER = "KISY", STAND_DATEI = "KISY_Stand.xlsx", SICHERUNG_ORDNER = "Sicherungen";
const XLSX_TYP = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const ORDNER_TYP = "application/vnd.google-apps.folder";

export const konfiguriert = () => !!GOOGLE_CLIENT_ID;
function lesen() { try { return JSON.parse(localStorage.getItem(KEY) || "{}"); } catch (e) { return {}; } }
function schreiben(o) { try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) { /* egal */ } }
export const angemeldet = () => { const g = lesen(); return !!g.token && g.bis > Date.now() + 60e3; };
export const warVerbunden = () => !!lesen().verbunden;
export const email = () => lesen().email || "";
// Rücksprungadresse: immer der Ordner der App (…/kisy/ bzw. localhost:8083/), nie …/index.html
const zurueckAdresse = () => location.origin + location.pathname.replace(/index\.html$/, "");

/* ---------- Anmeldung (Weiterleitung zu Google und zurück) ---------- */
export function anmelden(still = false) {
  const g = lesen();
  g.state = Array.from(crypto.getRandomValues(new Uint8Array(16)), b => b.toString(16).padStart(2, "0")).join("");
  g.seite = location.hash.startsWith("#/") ? location.hash : ""; // nach der Rückkehr wieder dieselbe Seite zeigen
  schreiben(g);
  const p = new URLSearchParams({ client_id: GOOGLE_CLIENT_ID, redirect_uri: zurueckAdresse(), response_type: "token", scope: SCOPES, include_granted_scopes: "false", state: g.state });
  if (still) p.set("prompt", "none");
  if (g.email) p.set("login_hint", g.email);
  location.assign("https://accounts.google.com/o/oauth2/v2/auth?" + p);
}
// Nach der Rückkehr von Google: Zugangsschlüssel aus der Adresse übernehmen. Ergebnis: null (keine Rückkehr), {ok} oder {fehler}
export function rueckkehrVerarbeiten() {
  if (!/access_token=|error=/.test(location.hash)) return null;
  const h = new URLSearchParams(location.hash.slice(1));
  const g = lesen();
  history.replaceState(null, "", location.pathname + location.search + (g.seite || ""));
  if (!g.state || h.get("state") !== g.state) return { fehler: "Anmeldung nicht bestätigt" };
  delete g.state; delete g.seite;
  if (h.get("error")) { schreiben(g); return { fehler: h.get("error") }; }
  g.token = h.get("access_token"); g.bis = Date.now() + (+h.get("expires_in") || 3600) * 1000; g.verbunden = true;
  schreiben(g); return { ok: true };
}
export function abmelden() {
  const g = lesen();
  if (g.token) fetch("https://oauth2.googleapis.com/revoke?token=" + encodeURIComponent(g.token), { method: "POST" }).catch(() => {});
  schreiben({});
}

async function api(url, opt = {}, erlaubt = []) {
  const g = lesen();
  if (!angemeldet()) throw Object.assign(new Error("Nicht mit Google verbunden"), { code: "anmelden" });
  const r = await fetch(url, { ...opt, headers: { ...(opt.headers || {}), Authorization: "Bearer " + g.token } });
  if (r.status === 401) { g.token = null; schreiben(g); throw Object.assign(new Error("Google-Anmeldung abgelaufen"), { code: "anmelden" }); }
  if (!r.ok && !erlaubt.includes(r.status)) throw Object.assign(new Error("Google meldet Fehler " + r.status), { code: "google", status: r.status });
  return r;
}
export async function kontoLaden() {
  const r = await api("https://www.googleapis.com/drive/v3/about?fields=user(emailAddress)");
  const e = (await r.json()).user?.emailAddress || ""; const g = lesen(); g.email = e; schreiben(g); return e;
}

/* ---------- Google Drive ---------- */
const q = s => encodeURIComponent(s);
async function suchen(bedingung) {
  const r = await api(`${DRIVE}?q=${q(bedingung + " and trashed=false")}&fields=files(id,name,modifiedTime)&orderBy=modifiedTime desc&spaces=drive`);
  return (await r.json()).files || [];
}
// Ordner suchen oder anlegen (gemerkt auf dem Gerät; neu = nicht dem gemerkten vertrauen)
async function ordner(name, eltern, merk, neu) {
  const g = lesen(); if (g[merk] && !neu) return g[merk];
  const f = await suchen(`name='${name}' and mimeType='${ORDNER_TYP}' and '${eltern}' in parents`);
  let id = f[0]?.id;
  if (!id) {
    const r = await api(DRIVE + "?fields=id", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name, mimeType: ORDNER_TYP, parents: [eltern] }) });
    id = (await r.json()).id;
  }
  const g2 = lesen(); g2[merk] = id; schreiben(g2); return id;
}
const hauptordner = neu => ordner(ORDNER, "root", "ordner", neu);
// Der KISY-Stand in Google Drive: { id, modifiedTime } oder null
export async function standInfo() {
  for (const neu of [false, true]) { // nichts gefunden: Ordner neu suchen (vielleicht wurde er verschoben oder gelöscht)
    const f = (await suchen(`name='${STAND_DATEI}' and '${await hauptordner(neu)}' in parents`))[0];
    if (f) return f;
  }
  return null;
}
export async function standLaden(id) { return (await api(`${DRIVE}/${id}?alt=media`)).arrayBuffer(); }
export async function standSpeichern(inhalt, id) {
  if (id) {
    const r = await api(`${UPLOAD}/${id}?uploadType=media&fields=id,modifiedTime`, { method: "PATCH", headers: { "Content-Type": XLSX_TYP }, body: inhalt }, [404]);
    if (r.ok) return r.json();
  }
  return anlegen(STAND_DATEI, inhalt, XLSX_TYP, await hauptordner(true));
}
async function anlegen(name, inhalt, typ, eltern) {
  const grenze = "kisy" + Date.now();
  const kopf = `--${grenze}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify({ name, parents: [eltern] })}\r\n--${grenze}\r\nContent-Type: ${typ}\r\n\r\n`;
  const body = new Blob([kopf, inhalt, `\r\n--${grenze}--`]);
  const r = await api(`${UPLOAD}?uploadType=multipart&fields=id,modifiedTime`, { method: "POST", headers: { "Content-Type": "multipart/related; boundary=" + grenze }, body });
  return r.json();
}
// Sicherungskopie in KISY › Sicherungen; eine vorhandene Datei mit gleichem Namen bleibt unangetastet (Rückgabe dann null)
export async function sicherungSpeichern(name, inhalt) {
  for (const neu of [false, true]) { // falls der Ordner inzwischen gelöscht wurde: einmal neu suchen/anlegen
    try {
      const o = await ordner(SICHERUNG_ORDNER, await hauptordner(neu), "sicherungen", neu);
      if ((await suchen(`name='${name}' and '${o}' in parents`)).length) return null;
      return await anlegen(name, inhalt, XLSX_TYP, o);
    } catch (e) { if (neu || e.code === "anmelden") throw e; }
  }
}
