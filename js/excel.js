// KISY-Stand als Excel-Datei schreiben und wieder lesen.
// Jedes Blatt hat eine ID-Spalte; damit hängen die Blätter zusammen. Namen in den Neben-Blättern sind nur zum Lesen da.
// Zeilen, die in Excel neu angelegt werden (ohne ID), bekommen beim Einlesen eine neue ID.
import { neueId, istDatum } from "./grundlagen.js";
import { neueDaten, normalisieren, kindVon, hcpLesen, gruppenLesen, FORMAT } from "./daten.js";

const EINV = "Einverständnis: ";
// Spalte „Gruppe“: mehrere Gruppen durch Komma getrennt („Minis, Fortgeschrittene“); der Name bleibt wie in Format 1,
// damit auch eine ältere KISY-Fassung die Gruppen beim Abgleich unverändert weitergibt.
const KINDER = [["ID", "id"], ["Vorname", "vorname"], ["Nachname", "nachname"], ["Geburtsdatum", "geburt", "d"], ["Gruppe", "gruppen", "g"],
  ["Dabei seit", "seit", "d"], ["Aktiv", "aktiv", "b"], ["Eltern", "eltern"], ["Telefon Eltern", "elternTel"], ["E-Mail Eltern", "elternMail"],
  ["Notfallkontakt", "notfall"], ["Notfall-Telefon", "notfallTel"], ["Hinweise (Allergien, Medikamente)", "hinweise"], ["Notiz", "notiz"]];

/* ---------- Schreiben ---------- */
// Excel zählt Tage ab dem 30.12.1899
const serial = s => (Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 864e5;
const DAT = Symbol("datum");
const dat = s => istDatum(s) ? { [DAT]: s } : "";
const leer = v => v == null ? "" : v;

function blatt(XLSX, kopf, zeilen, breiten) {
  const ws = XLSX.utils.aoa_to_sheet([kopf, ...zeilen.map(z => z.map(v => v && v[DAT] ? "" : leer(v)))]);
  // Datumszellen als echte Excel-Daten (sortierbar), ohne Zeitzonen-Verschiebung
  zeilen.forEach((z, r) => z.forEach((v, c) => {
    if (v && v[DAT]) ws[XLSX.utils.encode_cell({ r: r + 1, c })] = { t: "n", v: serial(v[DAT]), z: "dd.mm.yyyy" };
  }));
  ws["!cols"] = kopf.map((k, i) => ({ wch: (breiten && breiten[i]) || Math.max(10, Math.min(40, k.length + 2)) }));
  return ws;
}
const namen = (D, id) => { const k = kindVon(D, id); return k ? [k.vorname, k.nachname] : ["", ""]; };
const jaNein = b => b ? "ja" : "nein";

export function schreibeStand(XLSX, D, jetzt = new Date()) {
  const wb = XLSX.utils.book_new();
  const add = (name, kopf, zeilen, breiten) => XLSX.utils.book_append_sheet(wb, blatt(XLSX, kopf, zeilen, breiten), name);
  const einvKopf = D.einvArten.map(a => EINV + a.name);
  add("Kinder", KINDER.map(s => s[0]).concat(einvKopf), D.kinder.map(k => KINDER.map(([, f, art]) => art === "d" ? dat(k[f]) : art === "b" ? jaNein(k[f]) : art === "g" ? k[f].join(", ") : k[f])
    .concat(D.einvArten.map(a => { const v = k.einv[a.id]; return !v ? "" : istDatum(v) ? dat(v) : "ja"; }))), [12, 14, 16, 13, 12, 12, 7, 18, 16, 22, 18, 16, 30, 30]);
  add("Handicap", ["Kind-ID", "Vorname", "Nachname", "Datum", "Handicap"],
    D.kinder.flatMap(k => k.hcp.map(h => [k.id, k.vorname, k.nachname, dat(h.d), h.w < 0 ? "+" + String(-h.w).replace(".", ",") : h.w])));
  add("Training", ["ID", "Datum", "Gruppe", "Notizen", "Anwesend (Anzahl)"],
    D.trainings.slice().sort((a, b) => a.d.localeCompare(b.d)).map(x => [x.id, dat(x.d), x.gruppe, x.notiz, x.da.length]), [12, 12, 14, 60, 10]);
  add("Anwesenheit", ["Training-ID", "Datum", "Kind-ID", "Vorname", "Nachname"],
    D.trainings.slice().sort((a, b) => a.d.localeCompare(b.d)).flatMap(x => x.da.map(id => [x.id, dat(x.d), id, ...namen(D, id)])));
  add("Abzeichen", ["ID", "Abzeichen", "Ausgeblendet"], D.abzeichen.map(a => [a.id, a.name, a.aus ? "ja" : ""]), [16, 28, 12]);
  const aName = id => (D.abzeichen.find(a => a.id === id) || {}).name || "";
  add("Abzeichen erreicht", ["ID", "Datum", "Kind-ID", "Vorname", "Nachname", "Abzeichen-ID", "Abzeichen", "Notiz"],
    D.erreicht.slice().sort((a, b) => a.d.localeCompare(b.d)).map(e => [e.id, dat(e.d), e.kind, ...namen(D, e.kind), e.abz, aName(e.abz), e.notiz]));
  add("Turniere", ["ID", "Datum", "Turnier", "Ort", "Löcher", "Notiz"],
    D.turniere.slice().sort((a, b) => a.d.localeCompare(b.d)).map(x => [x.id, dat(x.d), x.name, x.ort, x.loecher ?? "", x.notiz]), [12, 12, 28, 18, 8, 30]);
  add("Turnier-Ergebnisse", ["Turnier-ID", "Datum", "Turnier", "Kind-ID", "Vorname", "Nachname", "Schläge", "Punkte", "Platz", "Notiz"],
    D.turniere.slice().sort((a, b) => a.d.localeCompare(b.d)).flatMap(x => x.tn.map(p => [x.id, dat(x.d), x.name, p.kind, ...namen(D, p.kind), p.schlaege ?? "", p.punkte ?? "", p.platz ?? "", p.notiz])));
  add("Listen", ["Art", "Wert", "ID"], [...D.gruppen.map(g => ["Gruppe", g, ""]), ...D.einvArten.map(a => ["Einverständnis", a.name, a.id])], [16, 28, 16]);
  const z = n => String(n).padStart(2, "0");
  add("Info", ["Feld", "Wert"], [["Programm", "KISY"], ["Format", FORMAT],
    ["Gespeichert am", `${z(jetzt.getDate())}.${z(jetzt.getMonth() + 1)}.${jetzt.getFullYear()} ${z(jetzt.getHours())}:${z(jetzt.getMinutes())}`],
    ["Letzte Sicherung in Google Drive", dat(D.sicherung || "")],
    ["Hinweis", "Die Blätter hängen über die ID-Spalten zusammen. IDs bitte nicht ändern."]], [16, 60]);
  return XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true });
}

/* ---------- Lesen ---------- */
const txt = v => v == null ? "" : v instanceof Date ? v.toISOString().slice(0, 10) : String(v).trim();
const ja = v => /^(ja|x|true|wahr|1)$/i.test(txt(v));
const nein = v => /^(nein|n|false|falsch|0)$/i.test(txt(v));
// Datum aus Excel (Zahl, Datum oder Text) -> "JJJJ-MM-TT"
export function datum(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return new Date(v.getTime() + 12 * 3600e3).toISOString().slice(0, 10);
  if (typeof v === "number") return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5).toISOString().slice(0, 10);
  const s = String(v).trim(); let m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return `${m[1]}-${m[2]}-${m[3]}`;
  if ((m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/))) return `${m[3].length === 2 ? "20" + m[3] : m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  return "";
}
const zahlOderNull = v => { if (typeof v === "number") return v; const s = txt(v).replace(",", "."); return s && !isNaN(+s) ? +s : null; };

// Liefert { daten, hinweise, format } – hinweise nennt Zeilen, die nicht zugeordnet werden konnten; format = Format der Datei.
export function leseStand(XLSX, daten) {
  const wb = XLSX.read(daten, { type: "array" });
  if (!wb.Sheets["Kinder"]) throw new Error("Das ist keine KISY-Datei (das Blatt „Kinder“ fehlt).");
  const rows = n => wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: "", raw: true }) : null;
  const D = neueDaten(), hinweise = [];
  const listen = rows("Listen");
  if (listen) {
    D.gruppen = listen.filter(z => txt(z["Art"]) === "Gruppe" && txt(z["Wert"])).map(z => txt(z["Wert"]));
    const ea = listen.filter(z => txt(z["Art"]) === "Einverständnis" && txt(z["Wert"])).map(z => ({ id: txt(z["ID"]) || neueId("e"), name: txt(z["Wert"]) }));
    D.einvArten = ea; // auch eine leere Liste ist gewollt
  }
  // Einverständnis-Spalten, die (noch) nicht in der Liste stehen, werden neue Arten
  const kopf = (XLSX.utils.sheet_to_json(wb.Sheets["Kinder"], { header: 1, range: 0 })[0] || []).map(txt);
  for (const s of kopf) if (s.startsWith(EINV) && s.length > EINV.length && !D.einvArten.some(a => a.name === s.slice(EINV.length).trim()))
    D.einvArten.push({ id: neueId("e"), name: s.slice(EINV.length).trim() });

  for (const z of rows("Kinder")) {
    if (!txt(z["Vorname"]) && !txt(z["Nachname"])) continue;
    const k = { id: txt(z["ID"]) || neueId("k"), einv: {}, hcp: [] };
    for (const [s, f, art] of KINDER) if (f !== "id") k[f] = art === "d" ? datum(z[s]) : art === "b" ? !nein(z[s]) : art === "g" ? gruppenLesen(txt(z[s])) : txt(z[s]);
    for (const a of D.einvArten) { const v = z[EINV + a.name]; const d = datum(v); if (d) k.einv[a.id] = d; else if (ja(v)) k.einv[a.id] = "ja"; }
    if (D.kinder.some(x => x.id === k.id)) { hinweise.push(`Kinder: ID ${k.id} doppelt – zweite Zeile hat eine neue ID bekommen`); k.id = neueId("k"); }
    D.kinder.push(k);
  }
  // Kind über ID finden, sonst über Vor- und Nachnamen
  const kindFinden = z => {
    const id = txt(z["Kind-ID"]);
    if (id && D.kinder.some(k => k.id === id)) return id;
    const v = txt(z["Vorname"]).toLowerCase(), n = txt(z["Nachname"]).toLowerCase();
    const L = D.kinder.filter(k => k.vorname.toLowerCase() === v && k.nachname.toLowerCase() === n);
    return L.length === 1 ? L[0].id : null;
  };
  const zeileName = z => [txt(z["Vorname"]), txt(z["Nachname"])].filter(Boolean).join(" ") || txt(z["Kind-ID"]) || "?";

  for (const z of rows("Handicap") || []) {
    const id = kindFinden(z), d = datum(z["Datum"]), w = hcpLesen(z["Handicap"]);
    if (!id) { hinweise.push(`Handicap: Kind „${zeileName(z)}“ nicht gefunden`); continue; }
    if (d && w != null) D.kinder.find(k => k.id === id).hcp.push({ d, w });
  }
  for (const z of rows("Training") || []) {
    const d = datum(z["Datum"]); if (!d) continue;
    D.trainings.push({ id: txt(z["ID"]) || neueId("t"), d, gruppe: txt(z["Gruppe"]), notiz: txt(z["Notizen"]), da: [] });
  }
  for (const z of rows("Anwesenheit") || []) {
    const id = kindFinden(z), d = datum(z["Datum"]);
    let x = D.trainings.find(t => t.id === txt(z["Training-ID"])) || (d && D.trainings.find(t => t.d === d));
    if (!x && d) { x = { id: neueId("t"), d, gruppe: "", notiz: "", da: [] }; D.trainings.push(x); }
    if (!id || !x) { hinweise.push(`Anwesenheit: „${zeileName(z)}“ am ${txt(z["Datum"]) || "?"} nicht zugeordnet`); continue; }
    if (!x.da.includes(id)) x.da.push(id);
  }
  // Blätter „Übungen“ und „Ergebnisse“ älterer Stände werden nicht mehr gelesen (Übungen gibt es seit Format 2 nicht mehr)
  const abz = rows("Abzeichen");
  if (abz) D.abzeichen = abz.filter(z => txt(z["Abzeichen"])).map(z => ({ id: txt(z["ID"]) || neueId("a"), name: txt(z["Abzeichen"]), aus: ja(z["Ausgeblendet"]) }));
  for (const z of rows("Abzeichen erreicht") || []) {
    const id = kindFinden(z), name = txt(z["Abzeichen"]);
    let a = D.abzeichen.find(x => x.id === txt(z["Abzeichen-ID"])) || D.abzeichen.find(x => x.name.toLowerCase() === name.toLowerCase());
    if (!a && name) { a = { id: neueId("a"), name, aus: false }; D.abzeichen.push(a); }
    if (!id || !a) { hinweise.push(`Abzeichen: „${zeileName(z)}“ nicht zugeordnet`); continue; }
    D.erreicht.push({ id: txt(z["ID"]) || neueId("z"), d: datum(z["Datum"]), kind: id, abz: a.id, notiz: txt(z["Notiz"]) });
  }
  for (const z of rows("Turniere") || []) {
    if (!txt(z["Turnier"]) && !datum(z["Datum"])) continue;
    D.turniere.push({ id: txt(z["ID"]) || neueId("w"), d: datum(z["Datum"]), name: txt(z["Turnier"]), ort: txt(z["Ort"]), loecher: zahlOderNull(z["Löcher"]), notiz: txt(z["Notiz"]), tn: [] });
  }
  for (const z of rows("Turnier-Ergebnisse") || []) {
    const id = kindFinden(z), d = datum(z["Datum"]);
    const x = D.turniere.find(t => t.id === txt(z["Turnier-ID"])) || D.turniere.find(t => t.name === txt(z["Turnier"]) && t.d === d);
    if (!id || !x) { hinweise.push(`Turnier-Ergebnisse: „${zeileName(z)}“ nicht zugeordnet`); continue; }
    if (!x.tn.some(p => p.kind === id)) x.tn.push({ kind: id, schlaege: zahlOderNull(z["Schläge"]), punkte: zahlOderNull(z["Punkte"]), platz: zahlOderNull(z["Platz"]), notiz: txt(z["Notiz"]) });
  }
  let format = 1;
  for (const z of rows("Info") || []) {
    if (txt(z["Feld"]) === "Letzte Sicherung in Google Drive") D.sicherung = datum(z["Wert"]);
    if (txt(z["Feld"]) === "Format") format = +txt(z["Wert"]) || 1;
  }
  D.version = format; // ältere Formate werden beim Normalisieren umgestellt (z. B. Startgruppen ergänzt)
  return { daten: normalisieren(D), hinweise, format };
}
