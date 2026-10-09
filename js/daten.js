// Datenmodell von KISY und alle Auswertungen (ohne Oberfläche, damit sie getestet werden können).
//
// D = {
//   version, kinder: [Kind], trainings: [Training], uebungen: [Übung], ergebnisse: [Ergebnis],
//   abzeichen: [Abzeichen], erreicht: [Erreicht], turniere: [Turnier], einvArten: [{id,name}], gruppen: [Text],
//   ungesichert, excelGesichert, geaendert
// }
// Kind      = { id, vorname, nachname, geburt, gruppe, seit, aktiv, notiz, eltern, elternTel, elternMail,
//               notfall, notfallTel, hinweise, einv: { artId: Datum | "ja" }, hcp: [{ d, w }] }   (w < 0 = Plus-Handicap)
// Training  = { id, d, gruppe ("" = alle), notiz, da: [Kind-ID] }
// Übung     = { id, name, einheit, max, besser: "hoch" | "niedrig", aus }
// Ergebnis  = { id, d, kind, uebung, wert, notiz }
// Abzeichen = { id, name, aus }          Erreicht = { id, d, kind, abz, notiz }
// Turnier   = { id, d, name, ort, loecher, notiz, tn: [{ kind, schlaege, punkte, platz, notiz }] }
import { neueId, vergleich, today, addMonths, zahl, zahlText, parseISO } from "./grundlagen.js";

export const FORMAT = 1;

export function startUebungen() {
  return [
    { id: "u-putten", name: "Putten", einheit: "Treffer", max: 10, besser: "hoch", aus: false },
    { id: "u-chippen", name: "Chippen", einheit: "Treffer", max: 10, besser: "hoch", aus: false },
    { id: "u-abschlag", name: "Abschlag-Weite", einheit: "m", max: null, besser: "hoch", aus: false },
  ];
}
export function startAbzeichen() {
  return [
    { id: "a-platzreife", name: "Platzreife", aus: false },
    { id: "a-dgv-bronze", name: "DGV-Abzeichen Bronze", aus: false },
    { id: "a-dgv-silber", name: "DGV-Abzeichen Silber", aus: false },
    { id: "a-dgv-gold", name: "DGV-Abzeichen Gold", aus: false },
  ];
}
export function startEinverstaendnisse() {
  return [{ id: "e-foto", name: "Foto-Erlaubnis" }, { id: "e-datenschutz", name: "Datenschutz-Erklärung" }];
}

export function neueDaten() {
  return {
    version: FORMAT, kinder: [], trainings: [], uebungen: startUebungen(), ergebnisse: [], abzeichen: startAbzeichen(),
    erreicht: [], turniere: [], einvArten: startEinverstaendnisse(), gruppen: [], sicherung: "", ungesichert: false, excelGesichert: "", geaendert: "",
  };
}

const t = v => v == null ? "" : String(v).trim();
export function neuesKind(werte = {}) {
  return normKind({ id: neueId("k"), seit: today(), aktiv: true, ...werte });
}
function normKind(k) {
  const n = { id: k.id || neueId("k"), vorname: t(k.vorname), nachname: t(k.nachname), geburt: t(k.geburt), gruppe: t(k.gruppe),
    seit: t(k.seit), aktiv: k.aktiv !== false, notiz: t(k.notiz), eltern: t(k.eltern), elternTel: t(k.elternTel), elternMail: t(k.elternMail),
    notfall: t(k.notfall), notfallTel: t(k.notfallTel), hinweise: t(k.hinweise), einv: { ...(k.einv || {}) },
    hcp: (k.hcp || []).filter(h => h && h.d && typeof h.w === "number").map(h => ({ d: h.d, w: h.w })) };
  for (const a in n.einv) if (!n.einv[a]) delete n.einv[a];
  n.hcp.sort((a, b) => a.d.localeCompare(b.d));
  return n;
}

// Fehlende Teile ergänzen (ältere oder unvollständige Stände), Verweise auf gelöschte Kinder entfernen.
export function normalisieren(D) {
  const N = { ...neueDaten(), ...(D || {}) };
  for (const f of ["kinder", "trainings", "ergebnisse", "erreicht", "turniere", "gruppen"]) if (!Array.isArray(N[f])) N[f] = [];
  for (const f of ["uebungen", "abzeichen", "einvArten"]) if (!Array.isArray(N[f])) N[f] = [];
  N.kinder = N.kinder.map(normKind);
  const ids = new Set(N.kinder.map(k => k.id));
  N.trainings = N.trainings.map(x => ({ id: x.id || neueId("t"), d: t(x.d), gruppe: t(x.gruppe), notiz: t(x.notiz), da: [...new Set((x.da || []).filter(id => ids.has(id)))] }));
  N.uebungen = N.uebungen.map(u => ({ id: u.id || neueId("u"), name: t(u.name), einheit: t(u.einheit), max: typeof u.max === "number" ? u.max : zahl(u.max), besser: u.besser === "niedrig" ? "niedrig" : "hoch", aus: !!u.aus }));
  N.abzeichen = N.abzeichen.map(a => ({ id: a.id || neueId("a"), name: t(a.name), aus: !!a.aus }));
  N.einvArten = N.einvArten.map(a => ({ id: a.id || neueId("e"), name: t(a.name) }));
  N.ergebnisse = N.ergebnisse.filter(e => ids.has(e.kind) && typeof e.wert === "number").map(e => ({ id: e.id || neueId("r"), d: t(e.d), kind: e.kind, uebung: e.uebung, wert: e.wert, notiz: t(e.notiz) }));
  N.erreicht = N.erreicht.filter(e => ids.has(e.kind)).map(e => ({ id: e.id || neueId("z"), d: t(e.d), kind: e.kind, abz: e.abz, notiz: t(e.notiz) }));
  N.turniere = N.turniere.map(x => ({ id: x.id || neueId("w"), d: t(x.d), name: t(x.name), ort: t(x.ort), loecher: typeof x.loecher === "number" ? x.loecher : zahl(x.loecher), notiz: t(x.notiz),
    tn: (x.tn || []).filter(p => ids.has(p.kind)).map(p => ({ kind: p.kind, schlaege: num(p.schlaege), punkte: num(p.punkte), platz: num(p.platz), notiz: t(p.notiz) })) }));
  // Gruppen: Liste plus alle, die bei Kindern oder Trainings vorkommen
  N.gruppen = [...new Set([...N.gruppen.map(t), ...N.kinder.map(k => k.gruppe), ...N.trainings.map(x => x.gruppe)].filter(Boolean))].sort(vergleich);
  N.version = FORMAT;
  return N;
}
const num = v => typeof v === "number" && isFinite(v) ? v : zahl(v);

/* ---------- Kinder ---------- */
export const kindName = k => k ? [k.vorname, k.nachname].filter(Boolean).join(" ") || "(ohne Namen)" : "(gelöscht)";
export const sortiereKinder = L => L.slice().sort((a, b) => vergleich(a.vorname, b.vorname) || vergleich(a.nachname, b.nachname));
export function kinderFiltern(D, { q = "", gruppe = "", inaktive = false } = {}) {
  const s = q.trim().toLowerCase();
  return sortiereKinder(D.kinder.filter(k => (inaktive || k.aktiv) && (!gruppe || k.gruppe === gruppe)
    && (!s || kindName(k).toLowerCase().includes(s) || k.gruppe.toLowerCase().includes(s))));
}
export const kindVon = (D, id) => D.kinder.find(k => k.id === id) || null;
// Alter in ganzen Jahren am Stichtag
export function alter(geburt, stichtag = today()) {
  if (!geburt) return null;
  const g = parseISO(geburt), s = parseISO(stichtag);
  let a = s.getFullYear() - g.getFullYear();
  if (s.getMonth() < g.getMonth() || (s.getMonth() === g.getMonth() && s.getDate() < g.getDate())) a--;
  return a >= 0 ? a : null;
}
// Kind mit allen Einträgen löschen (Anwesenheit, Ergebnisse, Abzeichen, Turniere)
export function kindLoeschen(D, id) {
  D.kinder = D.kinder.filter(k => k.id !== id);
  for (const x of D.trainings) x.da = x.da.filter(k => k !== id);
  D.ergebnisse = D.ergebnisse.filter(e => e.kind !== id);
  D.erreicht = D.erreicht.filter(e => e.kind !== id);
  for (const x of D.turniere) x.tn = x.tn.filter(p => p.kind !== id);
}
// Mehrere Kinder aus Text (ein Kind pro Zeile, "Vorname Nachname" oder "Nachname, Vorname")
export function namenLesen(text) {
  return String(text || "").split(/\r?\n/).map(z => z.trim()).filter(Boolean).map(z => {
    if (z.includes(",")) { const [nach, ...vor] = z.split(","); return { vorname: vor.join(",").trim(), nachname: nach.trim() }; }
    const teile = z.split(/\s+/); return { vorname: teile[0], nachname: teile.slice(1).join(" ") };
  });
}

/* ---------- Gruppen ---------- */
export function gruppeUmbenennen(D, alt, neu) {
  neu = t(neu); if (!neu || neu === alt) return false;
  for (const k of D.kinder) if (k.gruppe === alt) k.gruppe = neu;
  for (const x of D.trainings) if (x.gruppe === alt) x.gruppe = neu;
  D.gruppen = [...new Set(D.gruppen.map(g => g === alt ? neu : g))].sort(vergleich);
  return true;
}
export const gruppeBenutzt = (D, g) => D.kinder.some(k => k.gruppe === g) || D.trainings.some(x => x.gruppe === g);

/* ---------- Handicap ---------- */
// "36,4" -> 36.4, "+1,2" -> -1.2 (Plus-Handicap), "54" -> 54; ungültig -> null
export function hcpLesen(v) {
  if (typeof v === "number") return isFinite(v) ? v : null;
  const s = t(v); if (!s) return null;
  if (s.startsWith("+")) { const n = zahl(s.slice(1)); return n == null ? null : -n; }
  const n = zahl(s); return n == null || n < 0 ? null : n;
}
export const hcpText = w => w == null ? "–" : w < 0 ? "+" + zahlText(-w) : zahlText(w);
export const hcpAktuell = k => k && k.hcp.length ? k.hcp[k.hcp.length - 1] : null;

/* ---------- Zeitraum ---------- */
export const ZEITRAEUME = [["jahr", "Dieses Jahr"], ["vorjahr", "Letztes Jahr"], ["3m", "Letzte 3 Monate"], ["12m", "Letzte 12 Monate"], ["alle", "Alles"], ["frei", "Von – bis"]];
export function zeitraum(art, von = "", bis = "", heute = today()) {
  const j = +heute.slice(0, 4);
  switch (art) {
    case "jahr": return { von: `${j}-01-01`, bis: heute };
    case "vorjahr": return { von: `${j - 1}-01-01`, bis: `${j - 1}-12-31` };
    case "3m": return { von: addMonths(heute, -3), bis: heute };
    case "12m": return { von: addMonths(heute, -12), bis: heute };
    case "frei": return { von: von || "0000-01-01", bis: bis || "9999-12-31" };
    default: return { von: "0000-01-01", bis: "9999-12-31" };
  }
}
export const imZeitraum = (d, zr) => !!d && d >= zr.von && d <= zr.bis;

/* ---------- Training und Anwesenheit ---------- */
export const trainingsSortiert = D => D.trainings.slice().sort((a, b) => b.d.localeCompare(a.d) || vergleich(a.gruppe, b.gruppe));
export function neuesTraining(D, d = today(), gruppe = "") {
  const x = { id: neueId("t"), d, gruppe, notiz: "", da: [] };
  D.trainings.push(x); return x;
}
// Kinder, die für ein Training in Frage kommen (aktiv und passende Gruppe) – bereits abgehakte immer dabei
export function kinderFuerTraining(D, x) {
  const passt = k => k.aktiv && (!x.gruppe || k.gruppe === x.gruppe);
  return sortiereKinder(D.kinder.filter(k => passt(k) || x.da.includes(k.id)));
}
export function anwesenheitSetzen(x, kindId, da) {
  const ist = x.da.includes(kindId);
  if (da && !ist) x.da.push(kindId); else if (!da && ist) x.da = x.da.filter(k => k !== kindId);
}
// Wie oft war ein Kind da? Gezählt werden Trainings im Zeitraum ab „Dabei seit“, die für alle oder für seine
// Gruppe waren – und jedes Training, bei dem es tatsächlich da war (auch bei einer anderen Gruppe).
export function anwesenheitKind(D, k, zr) {
  const L = D.trainings.filter(x => imZeitraum(x.d, zr));
  let da = 0, von = 0, zuletzt = "";
  for (const x of L) {
    const war = x.da.includes(k.id);
    const zaehlt = war || ((!k.seit || x.d >= k.seit) && (!x.gruppe || x.gruppe === k.gruppe));
    if (zaehlt) von++;
    if (war) { da++; if (x.d > zuletzt) zuletzt = x.d; }
  }
  return { kind: k, da, von, quote: von ? da / von : null, zuletzt };
}
export function anwesenheit(D, zr, { gruppe = "", inaktive = false, sortierung = "quote" } = {}) {
  const L = kinderFiltern(D, { gruppe, inaktive }).map(k => anwesenheitKind(D, k, zr));
  if (sortierung === "quote") L.sort((a, b) => (b.quote ?? -1) - (a.quote ?? -1) || b.da - a.da || vergleich(kindName(a.kind), kindName(b.kind)));
  return L;
}
export const trainingsImZeitraum = (D, zr, gruppe = "") => D.trainings.filter(x => imZeitraum(x.d, zr) && (!gruppe || !x.gruppe || x.gruppe === gruppe)).length;

/* ---------- Übungen und Ergebnisse ---------- */
export const uebungVon = (D, id) => D.uebungen.find(u => u.id === id) || null;
export const aktiveUebungen = D => D.uebungen.filter(u => !u.aus);
export const einheitText = u => u ? (u.max != null ? `von ${zahlText(u.max)}` : u.einheit) : "";
export const wertText = (u, w) => w == null ? "–" : zahlText(w) + (u && u.max != null ? ` / ${zahlText(u.max)}` : u && u.einheit ? " " + u.einheit : "");
// Eingabe prüfen: Zahl, nicht negativ, höchstens Höchstwert. Liefert { wert } oder { fehler }.
export function wertPruefen(u, eingabe) {
  const s = t(eingabe); if (!s) return { wert: null };
  const w = zahl(s);
  if (w == null) return { fehler: "Bitte eine Zahl eingeben" };
  if (w < 0) return { fehler: "Bitte keine negative Zahl" };
  if (u && u.max != null && w > u.max) return { fehler: `Höchstens ${zahlText(u.max)}` };
  return { wert: w };
}
// Ergebnis eines Kindes für Übung + Tag setzen (eins pro Kind, Übung und Tag); null löscht es.
export function ergebnisSetzen(D, d, uebung, kind, wert) {
  const i = D.ergebnisse.findIndex(e => e.d === d && e.uebung === uebung && e.kind === kind);
  if (wert == null) { if (i >= 0) D.ergebnisse.splice(i, 1); return; }
  if (i >= 0) D.ergebnisse[i].wert = wert; else D.ergebnisse.push({ id: neueId("r"), d, kind, uebung, wert, notiz: "" });
}
export const ergebnisAm = (D, d, uebung, kind) => D.ergebnisse.find(e => e.d === d && e.uebung === uebung && e.kind === kind) || null;
export const ergebnisseVon = (D, kind, uebung) => D.ergebnisse.filter(e => e.kind === kind && (!uebung || e.uebung === uebung)).sort((a, b) => a.d.localeCompare(b.d));
export const besser = (u, a, b) => u && u.besser === "niedrig" ? a < b : a > b;
// Entwicklung eines Kindes bei einer Übung: erstes, letztes, bestes Ergebnis, Veränderung (positiv = besser geworden)
export function entwicklungKind(D, kind, u, zr = zeitraum("alle")) {
  const L = ergebnisseVon(D, kind, u.id).filter(e => imZeitraum(e.d, zr));
  if (!L.length) return null;
  let best = L[0];
  for (const e of L) if (besser(u, e.wert, best.wert)) best = e;
  const erstes = L[0], letztes = L[L.length - 1];
  const diff = letztes.wert - erstes.wert;
  return { anzahl: L.length, erstes, letztes, bestes: best, werte: L.map(e => ({ d: e.d, w: e.wert })), fortschritt: u.besser === "niedrig" ? -diff : diff,
    schnitt: L.reduce((s, e) => s + e.wert, 0) / L.length };
}
export function entwicklung(D, u, zr, { gruppe = "", inaktive = false } = {}) {
  return kinderFiltern(D, { gruppe, inaktive }).map(k => ({ kind: k, e: entwicklungKind(D, k.id, u, zr) }))
    .sort((a, b) => (!!b.e - !!a.e) || (b.e && a.e ? (besser(u, b.e.letztes.wert, a.e.letztes.wert) ? 1 : besser(u, a.e.letztes.wert, b.e.letztes.wert) ? -1 : 0) : 0) || vergleich(kindName(a.kind), kindName(b.kind)));
}
export const uebungBenutzt = (D, id) => D.ergebnisse.some(e => e.uebung === id);

/* ---------- Abzeichen ---------- */
export const abzeichenVon = (D, id) => D.abzeichen.find(a => a.id === id) || null;
export const abzeichenKind = (D, kind) => D.erreicht.filter(e => e.kind === kind).sort((a, b) => a.d.localeCompare(b.d));
export const hatAbzeichen = (D, kind, abz) => D.erreicht.some(e => e.kind === kind && e.abz === abz);
export function abzeichenEintragen(D, abz, d, kinder) {
  let n = 0;
  for (const kind of kinder) if (!hatAbzeichen(D, kind, abz)) { D.erreicht.push({ id: neueId("z"), d, kind, abz, notiz: "" }); n++; }
  return n;
}
// Übersicht je Abzeichen: wer hat es (mit Datum), wer noch nicht (aktive Kinder)
export function abzeichenUebersicht(D, { gruppe = "" } = {}) {
  const kinder = kinderFiltern(D, { gruppe });
  return D.abzeichen.filter(a => !a.aus).map(a => {
    const hat = D.erreicht.filter(e => e.abz === a.id).map(e => ({ e, kind: kindVon(D, e.kind) }))
      .filter(x => x.kind && (!gruppe || x.kind.gruppe === gruppe)).sort((x, y) => vergleich(kindName(x.kind), kindName(y.kind)));
    const ids = new Set(hat.map(x => x.kind.id));
    return { abz: a, hat, fehlt: kinder.filter(k => !ids.has(k.id)) };
  });
}
export const abzeichenBenutzt = (D, id) => D.erreicht.some(e => e.abz === id);

/* ---------- Turniere ---------- */
export const turniereSortiert = D => D.turniere.slice().sort((a, b) => b.d.localeCompare(a.d));
export function neuesTurnier(D, werte = {}) {
  const x = { id: neueId("w"), d: werte.d || today(), name: t(werte.name), ort: t(werte.ort), loecher: werte.loecher ?? null, notiz: "", tn: [] };
  D.turniere.push(x); return x;
}
export function turniereKind(D, kind, zr = zeitraum("alle")) {
  return turniereSortiert(D).filter(x => imZeitraum(x.d, zr)).map(x => ({ turnier: x, p: x.tn.find(p => p.kind === kind) })).filter(x => x.p);
}
export function turnierStatistik(D, zr, { gruppe = "" } = {}) {
  return kinderFiltern(D, { gruppe }).map(k => {
    const L = turniereKind(D, k.id, zr), plaetze = L.map(x => x.p.platz).filter(p => p != null);
    return { kind: k, anzahl: L.length, besterPlatz: plaetze.length ? Math.min(...plaetze) : null };
  }).filter(x => x.anzahl).sort((a, b) => b.anzahl - a.anzahl || (a.besterPlatz ?? 999) - (b.besterPlatz ?? 999));
}

/* ---------- Einverständnisse ---------- */
export const einvBenutzt = (D, id) => D.kinder.some(k => k.einv[id]);
export const einvFehlt = (D, k) => D.einvArten.filter(a => !k.einv[a.id]);

/* ---------- Zusammenfassung (z. B. vor dem Laden einer Excel-Datei) ---------- */
export function zaehlen(D) {
  return { kinder: D.kinder.length, trainings: D.trainings.length, ergebnisse: D.ergebnisse.length, abzeichen: D.erreicht.length, turniere: D.turniere.length };
}
export const zaehlenText = z => `${z.kinder} Kinder, ${z.trainings} Trainings, ${z.ergebnisse} Übungsergebnisse, ${z.abzeichen} Abzeichen, ${z.turniere} Turniere`;
