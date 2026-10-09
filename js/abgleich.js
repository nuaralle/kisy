// Abgleich zwischen Geräten über eine Datei in Google Drive (ohne Oberfläche, damit er getestet werden kann).
//
// Drei Stände werden verglichen („Dreiwege-Abgleich“):
//   basis = der Stand beim letzten Abgleich dieses Geräts, lokal = dieses Gerät jetzt, fern = Datei in Google Drive jetzt.
// Was nur auf einer Seite geändert wurde, wird übernommen; neue Einträge beider Seiten bleiben erhalten;
// Gelöschtes bleibt gelöscht. Wurde derselbe Eintrag auf beiden Seiten geändert, wird Feld für Feld zusammengeführt
// (Anwesenheit: beide Häkchen-Listen zusammen); nur wenn dasselbe Feld auf beiden Seiten anders ist, gilt dieses Gerät.
import { normalisieren, neueDaten, FORMAT } from "./daten.js";
import { schreibeStand, leseStand } from "./excel.js";

const LISTEN = ["kinder", "trainings", "uebungen", "ergebnisse", "abzeichen", "erreicht", "turniere", "einvArten"];
// Listen innerhalb eines Eintrags und ihr Schlüssel
const SCHLUESSEL = { hcp: "d", tn: "kind" };

/* ---------- Vergleichen ---------- */
// Einheitliche Textform: Schlüssel sortiert, Listen ohne Reihenfolge (Anwesenheit, Teilnehmer, IDs)
function kanon(v, feld) {
  if (Array.isArray(v)) {
    const k = SCHLUESSEL[feld] || "id";
    const L = v.map(x => kanon(x));
    return "[" + (v.length && typeof v[0] === "object" && v[0] && k in v[0]
      ? v.map((x, i) => [String(x[k]), L[i]]).sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0).map(x => x[1])
      : L.slice().sort()).join(",") + "]";
  }
  if (v && typeof v === "object") return "{" + Object.keys(v).sort().map(k => JSON.stringify(k) + ":" + kanon(v[k], k)).join(",") + "}";
  return JSON.stringify(v === undefined ? null : v);
}
const gleich = (a, b, feld) => kanon(a, feld) === kanon(b, feld);
// Nur der Inhalt (ohne Speicher- und Abgleich-Vermerke des Geräts)
export function inhalt(D) {
  const o = {}; for (const f of LISTEN) o[f] = D[f] || [];
  o.gruppen = (D.gruppen || []).slice().sort(); o.sicherung = D.sicherung || "";
  return kanon(o);
}
export const gleicherInhalt = (a, b) => inhalt(a) === inhalt(b);
// Hat das Gerät eigene Einträge (mehr als die Startlisten)?
export const hatEintraege = D => ["kinder", "trainings", "ergebnisse", "erreicht", "turniere"].some(f => (D[f] || []).length);

/* ---------- Zusammenführen ---------- */
function mengen(b = [], l = [], r = []) {
  const weg = new Set([...b.filter(x => !l.includes(x)), ...b.filter(x => !r.includes(x))]);
  return [...new Set([...l, ...r])].filter(x => !weg.has(x));
}
function feldweise(x, y, z, fernVorrang) {
  const o = {};
  for (const k of new Set([...Object.keys(y), ...Object.keys(z)])) {
    const a = x ? x[k] : undefined, l = y[k], r = z[k];
    if (gleich(l, r, k)) o[k] = l;
    else if (x && gleich(l, a, k)) o[k] = r;
    else if (x && gleich(r, a, k)) o[k] = l;
    else if (Array.isArray(l) && Array.isArray(r)) {
      const B = Array.isArray(a) ? a : [];
      o[k] = l.concat(r).every(e => e === null || typeof e !== "object")
        ? mengen(B, l, r) : liste(B, l, r, SCHLUESSEL[k] || "id", fernVorrang);
    }
    else if (l && r && typeof l === "object" && typeof r === "object") o[k] = feldweise(a && typeof a === "object" ? a : null, l, r, fernVorrang);
    else o[k] = fernVorrang ? r : l;
  }
  return o;
}
function liste(b, l, r, key, fernVorrang) {
  const B = new Map(b.map(e => [e[key], e])), Lm = new Map(l.map(e => [e[key], e])), Rm = new Map(r.map(e => [e[key], e]));
  const ids = [...new Set([...l.map(e => e[key]), ...r.map(e => e[key])])], aus = [];
  for (const id of ids) {
    const x = B.get(id), y = Lm.get(id), z = Rm.get(id);
    if (y && z) aus.push(gleich(y, z) ? y : x && gleich(y, x) ? z : x && gleich(z, x) ? y : feldweise(x, y, z, fernVorrang));
    else if (y) { if (!x || !gleich(y, x)) aus.push(y); }   // nur hier: neu oder hier geändert (dort gelöscht) → bleibt
    else if (z) { if (!x || !gleich(z, x)) aus.push(z); }
  }
  return aus;
}
// basis darf null sein (erster Abgleich dieses Geräts): dann wird alles vereinigt, bei Unterschieden gilt Google Drive.
export function zusammenfuehren(basis, lokal, fern) {
  const ohneBasis = !basis, B = basis || { gruppen: [] };
  const M = {};
  for (const f of LISTEN) M[f] = liste(B[f] || [], lokal[f] || [], fern[f] || [], "id", ohneBasis);
  M.gruppen = mengen(B.gruppen || [], lokal.gruppen || [], fern.gruppen || []);
  M.sicherung = [lokal.sicherung || "", fern.sicherung || ""].sort().pop();
  doppelteZusammenlegen(M, B);
  return normalisieren(M);
}
// Auf zwei Geräten unabhängig angelegte Doppel zusammenlegen:
// dasselbe Training (Tag + Gruppe), dasselbe Ergebnis (Tag + Übung + Kind), dasselbe Abzeichen beim Kind.
function doppelteZusammenlegen(M, B) {
  const inBasis = new Set((B.trainings || []).map(x => x.id)), schon = new Map(), raus = new Set();
  for (const x of M.trainings) {
    const k = x.d + "|" + x.gruppe, erstes = schon.get(k);
    if (!erstes) { schon.set(k, x); continue; }
    if (inBasis.has(x.id) && inBasis.has(erstes.id)) continue; // beide schon länger da: bewusst zwei Trainings
    erstes.da = [...new Set([...erstes.da, ...x.da])];
    if (x.notiz && !erstes.notiz.includes(x.notiz)) erstes.notiz = [erstes.notiz, x.notiz].filter(Boolean).join("\n");
    raus.add(x);
  }
  M.trainings = M.trainings.filter(x => !raus.has(x));
  const einmal = (L, schl) => { const s = new Set(); return L.filter(e => { const k = schl(e); if (s.has(k)) return false; s.add(k); return true; }); };
  M.ergebnisse = einmal(M.ergebnisse, e => e.d + "|" + e.uebung + "|" + e.kind);
  M.erreicht = einmal(M.erreicht, e => e.abz + "|" + e.kind);
}

/* ---------- Ein Abgleich-Durchgang ---------- */
// geraet = { daten, basis }   daten.drive = { id, modifiedTime } vom letzten Abgleich, daten.driveOffen = Änderungen seitdem,
//                             daten.ersetzen = Stand dieses Geräts soll den in Google Drive ersetzen (nach Excel laden)
// drive  = { info() -> {id, modifiedTime} | null, laden(id) -> ArrayBuffer, speichern(inhalt, id|null) -> {id, modifiedTime},
//            sichern?(inhalt, zusatz) }
// Ergebnis: { daten, basis, aktion: "nichts" | "hochgeladen" | "geladen" | "zusammengefuehrt" | "ersetzt" }
export async function abgleich(geraet, drive, XLSX) {
  const L = geraet.daten;
  const fertig = (D, meta, aktion) => {
    const N = { ...D, drive: { id: meta.id, modifiedTime: meta.modifiedTime }, driveOffen: false, ersetzen: false };
    return { daten: N, basis: nurInhalt(N), aktion };
  };
  for (let versuch = 0; versuch < 3; versuch++) {
    const meta = await drive.info();
    if (!meta) return fertig(L, await drive.speichern(schreibeStand(XLSX, L), null), "hochgeladen");
    const bekannt = L.drive && L.drive.id === meta.id && L.drive.modifiedTime === meta.modifiedTime;
    if (L.ersetzen) {
      if (drive.sichern) await drive.sichern(await drive.laden(meta.id), "vor_Ersetzen");
      return fertig(L, await drive.speichern(schreibeStand(XLSX, L), meta.id), "ersetzt");
    }
    if (bekannt) {
      if (!L.driveOffen) return { daten: L, basis: geraet.basis || nurInhalt(L), aktion: "nichts" };
      return fertig(L, await drive.speichern(schreibeStand(XLSX, L), meta.id), "hochgeladen");
    }
    const R = leseStand(XLSX, await drive.laden(meta.id)).daten;
    const M = !geraet.basis && !hatEintraege(L) ? R : zusammenfuehren(geraet.basis || null, L, R);
    const D = { ...M, ungesichert: L.ungesichert, excelGesichert: L.excelGesichert, geaendert: L.geaendert };
    if (gleicherInhalt(M, R)) return fertig(D, meta, "geladen");
    // Kurz vor dem Hochladen prüfen, ob in der Zwischenzeit ein anderes Gerät gespeichert hat – dann noch einmal
    const jetzt = await drive.info();
    if (jetzt && jetzt.modifiedTime !== meta.modifiedTime) continue;
    return fertig(D, await drive.speichern(schreibeStand(XLSX, M), meta.id), "zusammengefuehrt");
  }
  throw Object.assign(new Error("Google Drive ändert sich gerade ständig – später noch einmal"), { code: "unruhig" });
}
// Basis = nur der Inhalt (als einfacher Datensatz)
export function nurInhalt(D) {
  const o = { version: FORMAT }; for (const f of LISTEN) o[f] = structuredClone(D[f] || []);
  o.gruppen = (D.gruppen || []).slice(); o.sicherung = D.sicherung || "";
  return o;
}
export const leererStand = () => nurInhalt(neueDaten());
