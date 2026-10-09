// Oberfläche von KISY. Seiten werden über die Adresse gewählt (#/training, #/kind/<id> …),
// Knöpfe und Felder melden sich über data-a (Klick), data-c (Änderung), data-i (Eingabe) und data-f (Formular).
import { esc, fmtD, fmtDW, today, iso, addDays, parseISO, zahlText, zeitstempel, istDatum, vergleich } from "./grundlagen.js";
import * as M from "./daten.js";
import { schreibeStand, leseStand } from "./excel.js";
import * as Sp from "./speicher.js";
import { verlauf } from "./grafik.js";
import * as G from "./google.js";
import * as AG from "./abgleich.js";

const $ = s => document.querySelector(s);
const XLSX = window.XLSX;
const XLSX_TYP = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const IOS = /iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
const VERSION = "1.2";

let D = null;              // alle Daten (siehe daten.js)
let ORDNER = null;         // Infos zum Ordner „Daten“, wenn KISY am PC über den eigenen Server läuft
let SPEICHER_OK = true;
const SYNC = { laeuft: false, fehler: "", zuletzt: "" }; // Abgleich mit Google Drive
let syncTimer = null, NEU_ZEICHNEN = false;
// Einstellungen der Ansicht (nur für diese Sitzung)
const A = { zr: "jahr", von: "", bis: "", gruppe: "", q: "", inaktive: false, sort: "quote", heute: false };

/* ---------- kleine Helfer ---------- */
const h = id => esc(encodeURIComponent(id));            // ID in einer Adresse (#/kind/…)
const sel = (wert, ist) => wert === ist ? " selected" : "";
const name = k => esc(M.kindName(k));
const telHref = nr => "tel:" + String(nr).replace(/[^\d+]/g, "");
const prozent = q => q == null ? "–" : Math.round(q * 100) + " %";
const fmtDT = s => { const d = new Date(s || ""); return isNaN(d) ? "–" : `${fmtD(iso(d))} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`; };
const ZR = () => M.zeitraum(A.zr, A.von, A.bis);
const zrName = () => A.zr === "frei" ? `${A.von ? fmtD(A.von) : "Anfang"} – ${A.bis ? fmtD(A.bis) : "heute"}` : (M.ZEITRAEUME.find(z => z[0] === A.zr) || [, ""])[1];
const training = id => D.trainings.find(x => x.id === id);
const turnier = id => D.turniere.find(x => x.id === id);
const montag = d => addDays(d, -((parseISO(d).getDay() + 6) % 7));
const mitGoogle = () => G.konfiguriert() && G.warVerbunden();

function toast(text, dauer = 2600) {
  const el = $("#toast"); el.textContent = text; el.classList.add("an");
  clearTimeout(toast.t); toast.t = setTimeout(() => el.classList.remove("an"), dauer);
}
function persist(geaendert = true) {
  if (geaendert) { D.ungesichert = true; D.geaendert = new Date().toISOString(); D.driveOffen = true; D.aenderung = (D.aenderung || 0) + 1; planeAbgleich(); }
  SPEICHER_OK = Sp.speichern(D);
  statusZeigen();
}
function statusZeigen() {
  const el = $("#status"); if (!el || !D) return;
  let t = "Gespeichert", warn = false, tipp = "Alles auf diesem Gerät gespeichert";
  if (!SPEICHER_OK) { t = "⚠ Speichern auf dem Gerät fehlgeschlagen"; warn = true; }
  else if (mitGoogle()) {
    tipp = "Google Drive";
    if (!G.angemeldet()) { t = D.driveOffen ? "Nur auf diesem Gerät – neu verbinden" : "Google: neu verbinden"; warn = true; tipp = "Antippen, um wieder mit Google zu verbinden"; }
    else if (!navigator.onLine) { t = D.driveOffen ? "Offline – wird später gespeichert" : "Offline"; warn = !!D.driveOffen; }
    else if (SYNC.fehler) { t = SYNC.fehler; warn = true; }
    else t = SYNC.laeuft ? "Wird abgeglichen …" : D.driveOffen ? "Wird gespeichert …" : "In Google Drive gespeichert";
  }
  else if (D.ungesichert) { t = "Noch nicht als Excel gesichert"; warn = true; tipp = "Unter „Mehr“ als Excel speichern"; }
  el.textContent = t; el.classList.toggle("dirty", warn); el.title = tipp;
}
// Leiser Hinweis, wenn länger nicht als Excel gesichert wurde (nicht nötig, wenn Google Drive verbunden ist)
function sicherungsHinweis() {
  if (mitGoogle() || !D.ungesichert || !D.kinder.length) return "";
  const tage = D.excelGesichert ? (Date.now() - new Date(D.excelGesichert)) / 864e5 : 99;
  if (tage < 7) return "";
  return `<div class="hinweis">Denk an die Sicherung: ${D.excelGesichert ? "Die letzte Excel-Sicherung ist über eine Woche her." : "Es gibt noch keine Excel-Sicherung."}
    <button class="klein" data-a="excelSpeichern">Jetzt als Excel speichern</button></div>`;
}
function nichtGefunden(text, zurueck, zurueckText) {
  return `<a class="zurueck" href="${zurueck}">‹ ${zurueckText}</a><div class="leer"><p>${text}</p></div>`;
}
function geh(ziel) { if (location.hash === ziel) render(); else location.hash = ziel; }

/* ---------- Bausteine ---------- */
function gruppeWahl(wert = A.gruppe, c = "gruppe", text = "Alle Gruppen", extra = "") {
  if (!D.gruppen.length) return "";
  return `<label>Gruppe<select data-c="${c}"${extra}><option value="">${text}</option>${D.gruppen.map(g => `<option value="${esc(g)}"${sel(g, wert)}>${esc(g)}</option>`).join("")}</select></label>`;
}
function zeitraumWahl(mitGruppe = true) {
  return `<div class="row">
    <label>Zeitraum<select data-c="zr">${M.ZEITRAEUME.map(([k, t]) => `<option value="${k}"${sel(k, A.zr)}>${t}</option>`).join("")}</select></label>
    ${A.zr === "frei" ? `<label class="schmal">von<input type="date" data-c="zrVon" value="${esc(A.von)}"></label><label class="schmal">bis<input type="date" data-c="zrBis" value="${esc(A.bis)}"></label>` : ""}
    ${mitGruppe ? gruppeWahl() : ""}</div>`;
}
// Häkchen für die Gruppen eines Kindes (ein Kind kann in mehreren Gruppen sein)
function gruppenHaken(gewaehlt = []) {
  if (!D.gruppen.length) return `<p class="klein muted">Noch keine Gruppen angelegt (unter „Mehr“).</p>`;
  return `<fieldset class="gruppenhaken"><legend>Gruppen</legend>${D.gruppen.map(g => `<label class="chk"><input type="checkbox" name="gruppe" value="${esc(g)}"${gewaehlt.includes(g) ? " checked" : ""}> ${esc(g)}</label>`).join("")}</fieldset>`;
}
const gruppenKlein = k => k.gruppen.length && D.gruppen.length > 1 ? `<small>${esc(M.gruppenText(k))}</small>` : "";

/* ---------- Seiten ---------- */
const SEITEN = { training: seiteTraining, kinder: seiteKinder, kind: seiteKind, abzeichen: seiteAbzeichen, turniere: seiteTurniere, turnier: seiteTurnier, mehr: seiteMehr };
const REITER = { kind: "kinder", turnier: "turniere" };
function route() {
  const teile = location.hash.replace(/^#\/?/, "").split("/").map(s => { try { return decodeURIComponent(s); } catch (e) { return s; } });
  return { teil: SEITEN[teile[0]] ? teile[0] : "training", rest: teile.slice(1) };
}
function render() {
  const r = route();
  const reiter = REITER[r.teil] || r.teil;
  document.querySelectorAll(".tabs a").forEach(a => a.dataset.t === reiter ? a.setAttribute("aria-current", "page") : a.removeAttribute("aria-current"));
  $("#main").innerHTML = SEITEN[r.teil](...r.rest);
  statusZeigen();
}

/* ----- Training ----- */
function umschalterTraining(aktiv) {
  return `<nav class="umschalter" aria-label="Ansicht"><a href="#/training"${aktiv === "tage" ? ' aria-current="page"' : ""}>Trainingstage</a><a href="#/training/uebersicht"${aktiv === "uebersicht" ? ' aria-current="page"' : ""}>Wer war wie oft da?</a></nav>`;
}
function seiteTraining(unter) {
  if (unter === "uebersicht") return seiteAnwesenheit();
  if (unter) return seiteTrainingTag(unter);
  const L = M.trainingsSortiert(D);
  return `<div class="kopf"><h2>Training</h2><button class="pri" data-a="trainingHeute">＋ Training heute</button></div>
  ${heuteWahl()}
  ${umschalterTraining("tage")}
  ${sicherungsHinweis()}
  ${!D.kinder.length && ORDNER && ORDNER.stand ? `<div class="hinweis gut">Im Ordner „Daten“ liegt ein gespeicherter Stand (${esc(ORDNER.stand)}). <button class="klein" data-a="ordnerLaden">Diesen Stand laden</button></div>` : ""}
  ${G.konfiguriert() && !G.warVerbunden() ? `<div class="hinweis gut">Mit Google Drive verbinden, damit iPhone und PC denselben Stand haben. <button class="klein" data-a="gVerbinden">Mit Google verbinden</button></div>` : ""}
  ${!D.kinder.length ? `<div class="hinweis">Lege zuerst unter <a href="#/kinder">Kinder</a> die Kinder an – dann kannst du hier abhaken, wer da war.</div>` : ""}
  <details class="karte"><summary>Training an einem anderen Tag anlegen</summary>
    <form data-f="trainingNeu" class="row"><label class="schmal">Datum<input type="date" name="d" value="${today()}" required></label>
    ${D.gruppen.length ? `<label>Gruppe<select name="gruppe"><option value="">Alle Kinder</option>${D.gruppen.map(g => `<option value="${esc(g)}">${esc(g)}</option>`).join("")}</select></label>` : ""}
    <button type="submit">Anlegen</button></form></details>
  ${L.length ? `<ul class="liste">${L.map(x => `<li><a class="zeile" href="#/training/${h(x.id)}"><span><b>${fmtDW(x.d)}</b>
    <small>${[x.gruppe || (D.gruppen.length ? "Alle Kinder" : ""), (x.notiz || "").split("\n")[0]].filter(Boolean).map(esc).join(" · ") || "&nbsp;"}</small></span>
    <span class="re"><b>${x.da.length}</b> da</span></a></li>`).join("")}</ul>` : `<div class="leer"><p>Noch keine Trainings eingetragen. Tippe auf „Training heute“.</p></div>`}`;
}
// „Training heute“: für welche Gruppe? (vorhandenes Training des Tages wird geöffnet)
function heuteWahl() {
  if (!A.heute || !D.gruppen.length) return "";
  const da = g => D.trainings.some(x => x.d === today() && x.gruppe === g);
  const knopf = (g, text, n) => `<button data-a="heuteGruppe" data-g="${esc(g)}"><b>${text}</b><small>${da(g) ? "schon angelegt – öffnen" : n + " Kinder"}</small></button>`;
  return `<div class="karte"><h3>Training heute – für welche Gruppe?</h3><div class="gruppenwahl">
    ${D.gruppen.map(g => knopf(g, esc(g), M.gruppeAnzahl(D, g))).join("")}${knopf("", "Alle Kinder", D.kinder.filter(k => k.aktiv).length)}</div>
    <button class="leise klein" data-a="heuteZu" style="margin-top:.5rem">Abbrechen</button></div>`;
}
const zaehlerText = x => { const L = M.kinderFuerTraining(D, x); return `<b>${x.da.length}</b> von ${L.length} da`; };
function anwKnopf(x, k) {
  const da = x.da.includes(k.id);
  return `<li><button type="button" class="anw" data-a="anw" data-id="${esc(x.id)}" data-k="${esc(k.id)}" aria-pressed="${da}"><span class="haken"></span>
    <span class="n">${name(k)}${!x.gruppe ? gruppenKlein(k) : ""}</span></button></li>`;
}
function seiteTrainingTag(id) {
  const x = training(id);
  if (!x) return nichtGefunden("Dieses Training gibt es nicht (mehr).", "#/training", "Alle Trainings");
  const L = M.kinderFuerTraining(D, x), andere = M.sortiereKinder(D.kinder.filter(k => k.aktiv && !L.includes(k)));
  return `<a class="zurueck" href="#/training">‹ Alle Trainings</a>
  <div class="kopf"><h2>${fmtDW(x.d)}</h2><div class="zaehler" id="zaehler">${zaehlerText(x)}</div></div>
  <div class="karte"><div class="row"><label class="schmal">Datum<input type="date" value="${esc(x.d)}" data-c="tDatum" data-id="${esc(x.id)}"></label>
    ${gruppeWahl(x.gruppe, "tGruppe", "Alle Kinder", ` data-id="${esc(x.id)}"`)}</div></div>
  <div class="kopf"><h3>Wer ist da?</h3><div class="knoepfe"><button class="klein" data-a="alleDa" data-id="${esc(x.id)}">Alle da</button><button class="klein leise" data-a="keinerDa" data-id="${esc(x.id)}">Keiner</button></div></div>
  ${L.length ? `<ul class="anwliste">${L.map(k => anwKnopf(x, k)).join("")}</ul>` : `<p class="muted">${D.kinder.length ? "Keine aktiven Kinder in dieser Gruppe." : `Noch keine Kinder angelegt – <a href="#/kinder">jetzt anlegen</a>.`}</p>`}
  ${andere.length ? `<details><summary>Kinder aus anderen Gruppen (${andere.length})</summary><ul class="anwliste">${andere.map(k => anwKnopf(x, k)).join("")}</ul></details>` : ""}
  <div class="karte" style="margin-top:1rem"><label>Notizen zum Training (Inhalt, Wetter, Besonderes)<textarea data-i="tNotiz" data-id="${esc(x.id)}" rows="4" placeholder="z. B. Putten 2 m, Chippen ins Ziel; sonnig, 14 °C">${esc(x.notiz)}</textarea></label></div>
  <div class="knoepfe"><button class="gefahr" data-a="trainingLoeschen" data-id="${esc(x.id)}">Training löschen</button></div>`;
}
function seiteAnwesenheit() {
  const zr = ZR(), L = M.anwesenheit(D, zr, { gruppe: A.gruppe, sortierung: A.sort }), n = M.trainingsImZeitraum(D, zr, A.gruppe);
  return `<div class="kopf"><h2>Training</h2></div>
  ${umschalterTraining("uebersicht")}
  <div class="karte">${zeitraumWahl()}<div class="row" style="margin-top:.6rem"><label>Sortierung<select data-c="sort"><option value="quote"${sel("quote", A.sort)}>nach Anwesenheit</option><option value="name"${sel("name", A.sort)}>nach Name</option></select></label></div></div>
  <p><b>${n}</b> Training${n === 1 ? "" : "s"} im Zeitraum (${esc(zrName())}${A.gruppe ? ", Gruppe " + esc(A.gruppe) : ""}).</p>
  ${L.length ? `<ul class="liste">${L.map(a => `<li><a class="zeile" href="#/kind/${h(a.kind.id)}"><span><b>${name(a.kind)}</b>
    <small>${a.zuletzt ? "zuletzt da am " + fmtD(a.zuletzt) : "im Zeitraum nicht da"}</small></span>
    <span class="re"><b>${a.da}</b> von ${a.von} · ${prozent(a.quote)}</span>
    <span class="balken" aria-hidden="true"><i style="width:${Math.round((a.quote || 0) * 100)}%"></i></span></a></li>`).join("")}</ul>`
    : `<div class="leer"><p>Noch keine Kinder angelegt.</p></div>`}
  <p class="klein muted">Gezählt werden die Trainings ab „Dabei seit“ des Kindes, die für alle oder für eine seiner Gruppen waren – und jedes Training, bei dem es da war.</p>`;
}

/* ----- Kinder ----- */
function kinderZeilen() {
  const L = M.kinderFiltern(D, { q: A.q, gruppe: A.gruppe, inaktive: A.inaktive }), zr = M.zeitraum("jahr");
  if (!L.length) return `<p class="muted">Kein Kind passt zur Suche.</p>`;
  return `<ul class="liste">${L.map(k => {
    const a = M.anwesenheitKind(D, k, zr), hc = M.hcpAktuell(k), al = M.alter(k.geburt), fehlt = M.einvFehlt(D, k).length;
    const info = [k.gruppen.length && esc(M.gruppenText(k)), al != null && al + " Jahre", hc && "Hcp " + M.hcpText(hc.w)].filter(Boolean).join(" · ");
    return `<li><a class="zeile" href="#/kind/${h(k.id)}"><span><b>${name(k)}</b>
      ${!k.aktiv ? ' <span class="marker">inaktiv</span>' : ""}${k.hinweise ? ' <span class="marker orange" title="Hinweis (z. B. Allergie)">⚠ Hinweis</span>' : ""}${fehlt ? ` <span class="marker orange">${fehlt === 1 ? "1 Einverständnis fehlt" : fehlt + " Einverständnisse fehlen"}</span>` : ""}
      <small>${info || "&nbsp;"}</small></span>
      <span class="re">${a.von ? `<b>${prozent(a.quote)}</b><br>da` : ""}</span></a></li>`;
  }).join("")}</ul>`;
}
function seiteKinder(unter) {
  if (unter === "mehrere") return seiteMehrere();
  return `<div class="kopf"><h2>Kinder</h2><div class="knoepfe"><a class="knopf pri" href="#/kind/neu">＋ Kind</a><a class="knopf" href="#/kinder/mehrere">Mehrere anlegen</a></div></div>
  ${D.kinder.length ? `<div class="karte"><div class="row"><label>Suchen<input type="search" data-i="q" value="${esc(A.q)}" placeholder="Name oder Gruppe" autocomplete="off"></label>
    ${gruppeWahl()}<label class="chk"><input type="checkbox" data-c="inaktive"${A.inaktive ? " checked" : ""}> auch inaktive</label></div></div>
    <div id="kliste">${kinderZeilen()}</div><p class="klein muted">Prozent = Anwesenheit beim Training in diesem Jahr.</p>`
    : `<div class="leer"><p>Noch keine Kinder eingetragen.</p><div class="knoepfe"><a class="knopf pri" href="#/kind/neu">＋ Erstes Kind anlegen</a><a class="knopf" href="#/kinder/mehrere">Mehrere auf einmal</a></div></div>`}`;
}
function seiteMehrere() {
  return `<a class="zurueck" href="#/kinder">‹ Alle Kinder</a><h2>Mehrere Kinder anlegen</h2>
  <form class="karte" data-f="mehrere" autocomplete="off" style="margin-top:1rem">
    <label>Ein Kind pro Zeile („Vorname Nachname“)<textarea name="namen" rows="8" required placeholder="Vorname Nachname&#10;Vorname Nachname"></textarea></label>
    <div style="margin-top:.6rem">${gruppenHaken()}<p class="klein muted" style="margin:.2rem 0 0">Die angekreuzten Gruppen bekommen alle Kinder aus der Liste.</p></div>
    <div class="row" style="margin-top:.6rem"><label class="schmal">Dabei seit<input type="date" name="seit" value="${today()}"></label></div>
    <div class="knoepfe" style="margin-top:1rem"><button class="pri" type="submit">Anlegen</button><a class="knopf" href="#/kinder">Abbrechen</a></div></form>`;
}
function seiteKind(id, modus) {
  if (id === "neu") return kindFormular(null);
  const k = M.kindVon(D, id);
  if (!k) return nichtGefunden("Dieses Kind gibt es nicht (mehr).", "#/kinder", "Alle Kinder");
  if (modus === "bearbeiten") return kindFormular(k);
  const zr = M.zeitraum("jahr"), a = M.anwesenheitKind(D, k, zr), hc = M.hcpAktuell(k), al = M.alter(k.geburt);
  const abz = M.abzeichenKind(D, k.id), tur = M.turniereKind(D, k.id);
  const offen = D.abzeichen.filter(x => !x.aus && !M.hatAbzeichen(D, k.id, x.id));
  const kopf = [k.gruppen.length && esc(M.gruppenText(k)), al != null && `${al} Jahre (geb. ${fmtD(k.geburt)})`, k.seit && `dabei seit ${fmtD(k.seit)}`].filter(Boolean).join(" · ");
  const tel = nr => nr ? `<a class="tel" href="${esc(telHref(nr))}">${esc(nr)}</a>` : "";
  return `<a class="zurueck" href="#/kinder">‹ Alle Kinder</a>
  <div class="kopf"><div><h2>${name(k)}${!k.aktiv ? ' <span class="marker">inaktiv</span>' : ""}</h2><p class="muted" style="margin:.2rem 0 0">${kopf}</p></div>
    <a class="knopf" href="#/kind/${h(k.id)}/bearbeiten">Bearbeiten</a></div>
  ${k.hinweise ? `<p class="achtung">⚠ ${esc(k.hinweise)}</p>` : ""}
  <div class="kacheln">
    <div><b>${a.da} / ${a.von}</b><span>Trainings dieses Jahr (${prozent(a.quote)})</span></div>
    <div><b>${hc ? M.hcpText(hc.w) : "–"}</b><span>Handicap${hc ? " (" + fmtD(hc.d) + ")" : ""}</span></div>
    <div><b>${abz.length}</b><span>Abzeichen</span></div>
    <div><b>${tur.length}</b><span>Turniere</span></div>
  </div>
  <section class="karte"><h3>Eltern &amp; Notfall</h3>
    ${k.eltern || k.elternTel || k.elternMail || k.notfall || k.notfallTel ? `<dl class="fakten">
      ${k.eltern || k.elternTel ? `<dt>Eltern</dt><dd>${esc(k.eltern)} ${tel(k.elternTel)}</dd>` : ""}
      ${k.elternMail ? `<dt>E-Mail</dt><dd><a href="mailto:${esc(k.elternMail)}">${esc(k.elternMail)}</a></dd>` : ""}
      ${k.notfall || k.notfallTel ? `<dt>Notfall</dt><dd>${esc(k.notfall)} ${tel(k.notfallTel)}</dd>` : ""}</dl>` : `<p class="muted">Keine Angaben. <a href="#/kind/${h(k.id)}/bearbeiten">Eintragen</a></p>`}
  </section>
  <section class="karte"><h3>Einverständnisse</h3>
    ${D.einvArten.length ? `<ul class="namen">${D.einvArten.map(e => { const v = k.einv[e.id]; return `<li><span>${esc(e.name)}</span><span class="${v ? "ok" : "warn"}">${v ? "✓ " + (istDatum(v) ? "seit " + fmtD(v) : "ja") : "fehlt"}</span></li>`; }).join("")}</ul>` : `<p class="muted">Keine Arten angelegt (unter „Mehr“).</p>`}
  </section>
  <section class="karte"><h3>Handicap ${k.hcp.length > 1 ? verlauf(k.hcp, { breite: 140 }) : ""}</h3>
    ${k.hcp.length ? `<ul class="namen">${k.hcp.slice().reverse().map(x => `<li><span>${fmtD(x.d)}</span><span><b>${M.hcpText(x.w)}</b> <button class="klein leise" data-a="hcpWeg" data-id="${esc(k.id)}" data-d="${esc(x.d)}" aria-label="Eintrag vom ${fmtD(x.d)} löschen">✕</button></span></li>`).join("")}</ul>` : `<p class="muted">Noch kein Handicap eingetragen.</p>`}
    <form data-f="hcp" data-id="${esc(k.id)}" class="row" style="margin-top:.6rem" autocomplete="off"><label class="schmal">Datum<input type="date" name="d" value="${today()}" required></label>
      <label class="schmal">Handicap<input name="w" inputmode="decimal" placeholder="z. B. 36,4" required></label><button type="submit">Eintragen</button></form>
    <p class="klein muted">Plus-Handicap mit „+“ davor eingeben (z. B. +1,2).</p>
  </section>
  <section class="karte"><h3>Abzeichen</h3>
    ${abz.length ? `<ul class="namen">${abz.map(e => `<li><span>${esc((M.abzeichenVon(D, e.abz) || {}).name || "?")}</span><span>${fmtD(e.d)} <button class="klein leise" data-a="erreichtWeg" data-id="${esc(e.id)}" aria-label="Abzeichen löschen">✕</button></span></li>`).join("")}</ul>` : `<p class="muted">Noch keine Abzeichen.</p>`}
    ${offen.length ? `<form data-f="kindAbz" data-id="${esc(k.id)}" class="row" style="margin-top:.6rem"><label>Geschafft<select name="abz">${offen.map(x => `<option value="${esc(x.id)}">${esc(x.name)}</option>`).join("")}</select></label>
      <label class="schmal">am<input type="date" name="d" value="${today()}" required></label><button type="submit">Eintragen</button></form>` : ""}
  </section>
  <section class="karte"><h3>Turniere</h3>
    ${tur.length ? `<ul class="namen">${tur.map(({ turnier: x, p }) => `<li><a href="#/turnier/${h(x.id)}">${fmtD(x.d)} ${esc(x.name || "Turnier")}</a><span>${turnierWerte(p)}</span></li>`).join("")}</ul>` : `<p class="muted">Noch keine Turniere.</p>`}
  </section>
  <section class="karte"><h3>Training</h3><p style="margin:0">Dieses Jahr <b>${a.da}</b> von ${a.von} Trainings da (${prozent(a.quote)})${a.zuletzt ? `, zuletzt am ${fmtD(a.zuletzt)}` : ""}.</p></section>
  ${k.notiz ? `<section class="karte"><h3>Notiz</h3><p style="margin:0;white-space:pre-wrap">${esc(k.notiz)}</p></section>` : ""}`;
}
const turnierWerte = p => [p.schlaege != null && `${zahlText(p.schlaege)} Schläge`, p.punkte != null && `${zahlText(p.punkte)} Punkte`, p.platz != null && `<b>Platz ${zahlText(p.platz)}</b>`].filter(Boolean).join(" · ") || "–";

function kindFormular(k) {
  const neu = !k; k = k || M.neuesKind();
  const f = (feld, text, art = "text", extra = "") => `<label>${text}<input type="${art}" name="${feld}" value="${esc(k[feld])}" autocomplete="off"${extra}></label>`;
  return `<a class="zurueck" href="${neu ? "#/kinder" : "#/kind/" + h(k.id)}">‹ ${neu ? "Alle Kinder" : "Zurück"}</a>
  <h2>${neu ? "Neues Kind" : name(k) + " bearbeiten"}</h2>
  <form data-f="kind" data-id="${neu ? "" : esc(k.id)}" autocomplete="off" style="margin-top:1rem">
  <fieldset class="karte"><h3>Kind</h3><div class="grid2">
    ${f("vorname", "Vorname *", "text", " required")}${f("nachname", "Nachname")}
    ${f("geburt", "Geburtsdatum", "date")}${f("seit", "Dabei seit", "date")}
    <label class="chk"><input type="checkbox" name="aktiv"${k.aktiv ? " checked" : ""}> aktiv (kommt zum Training)</label>
  </div><div style="margin-top:.7rem">${gruppenHaken(k.gruppen)}</div>
  <label style="margin-top:.7rem">Notiz<textarea name="notiz" rows="3">${esc(k.notiz)}</textarea></label></fieldset>
  <fieldset class="karte"><h3>Eltern &amp; Notfall</h3><div class="grid2">
    ${f("eltern", "Eltern (Name)")}${f("elternTel", "Telefon Eltern", "tel")}
    ${f("elternMail", "E-Mail Eltern", "email")}<span></span>
    ${f("notfall", "Notfallkontakt (Name)")}${f("notfallTel", "Notfall-Telefon", "tel")}
  </div><label style="margin-top:.7rem">Hinweise (Allergien, Medikamente, Besonderes)<textarea name="hinweise" rows="2">${esc(k.hinweise)}</textarea></label></fieldset>
  <fieldset class="karte"><h3>Einverständnisse</h3>
    ${D.einvArten.length ? D.einvArten.map(e => { const v = k.einv[e.id] || ""; return `<div class="row" style="align-items:center;margin-bottom:.4rem">
      <label class="chk" style="flex:1 1 12rem"><input type="checkbox" name="einv_${esc(e.id)}" data-c="einvHaken"${v ? " checked" : ""}> ${esc(e.name)}</label>
      <label class="schmal">Datum<input type="date" name="einvd_${esc(e.id)}" value="${istDatum(v) ? v : ""}"></label></div>`; }).join("") : `<p class="muted">Keine Arten angelegt (unter „Mehr“).</p>`}
  </fieldset>
  <div class="knoepfe"><button class="pri" type="submit">Speichern</button><a class="knopf" href="${neu ? "#/kinder" : "#/kind/" + h(k.id)}">Abbrechen</a>
    ${neu ? "" : `<button type="button" class="gefahr" data-a="kindLoeschen" data-id="${esc(k.id)}">Kind löschen</button>`}</div>
  </form>`;
}

/* ----- Abzeichen ----- */
function seiteAbzeichen(id) {
  if (id) return seiteAbzEintragen(id);
  const U = M.abzeichenUebersicht(D, { gruppe: A.gruppe });
  return `<div class="kopf"><h2>Abzeichen</h2><a class="knopf" href="#/mehr/listen">Liste bearbeiten</a></div>
  ${D.gruppen.length ? `<div class="karte"><div class="row">${gruppeWahl()}</div></div>` : ""}
  ${U.length ? `<div class="abzliste">${U.map(x => `<section class="karte"><h3>${esc(x.abz.name)} <span class="marker gruen">${x.hat.length}</span><a class="knopf klein rechts" href="#/abzeichen/${h(x.abz.id)}">Eintragen</a></h3>
    ${x.hat.length ? `<ul class="namen">${x.hat.map(y => `<li><a href="#/kind/${h(y.kind.id)}">${name(y.kind)}</a><span>${fmtD(y.e.d)}</span></li>`).join("")}</ul>` : `<p class="muted" style="margin:0">Noch niemand.</p>`}
    ${x.fehlt.length && x.hat.length ? `<details><summary>Noch nicht geschafft (${x.fehlt.length})</summary><p class="fehlt">${x.fehlt.map(name).join(", ")}</p></details>` : ""}</section>`).join("")}</div>`
    : `<div class="leer"><p>Es sind keine Abzeichen angelegt.</p><a class="knopf pri" href="#/mehr/listen">Abzeichen anlegen</a></div>`}`;
}
function seiteAbzEintragen(id) {
  const a = M.abzeichenVon(D, id);
  if (!a) return nichtGefunden("Dieses Abzeichen gibt es nicht (mehr).", "#/abzeichen", "Abzeichen");
  const kinder = M.kinderFiltern(D, { gruppe: A.gruppe }), offen = kinder.filter(k => !M.hatAbzeichen(D, k.id, a.id)), hat = kinder.filter(k => M.hatAbzeichen(D, k.id, a.id));
  return `<a class="zurueck" href="#/abzeichen">‹ Abzeichen</a><h2>${esc(a.name)} eintragen</h2>
  <form data-f="abzEintragen" data-id="${esc(a.id)}" style="margin-top:1rem">
    <div class="karte"><div class="row"><label class="schmal">Geschafft am<input type="date" name="d" value="${today()}" required></label>${gruppeWahl()}</div></div>
    <h3 style="margin:0 0 .5rem">Wer hat es geschafft?</h3>
    ${offen.length ? `<ul class="liste">${offen.map(k => `<li><label class="chk auswahl"><input type="checkbox" name="k" value="${esc(k.id)}"> ${name(k)}${!A.gruppe ? gruppenKlein(k) : ""}</label></li>`).join("")}</ul>`
      : `<p class="muted">Alle ${A.gruppe ? "Kinder dieser Gruppe" : "aktiven Kinder"} haben dieses Abzeichen schon.</p>`}
    <div class="knoepfe" style="margin-top:1rem"><button class="pri" type="submit"${offen.length ? "" : " disabled"}>Speichern</button><a class="knopf" href="#/abzeichen">Abbrechen</a></div>
  </form>
  ${hat.length ? `<p class="fehlt" style="margin-top:1rem">Schon geschafft: ${hat.map(name).join(", ")}</p>` : ""}`;
}

/* ----- Turniere ----- */
function seiteTurniere() {
  const L = M.turniereSortiert(D), S = M.turnierStatistik(D, ZR(), { gruppe: A.gruppe });
  return `<div class="kopf"><h2>Turniere</h2></div>
  <details class="karte"${L.length ? "" : " open"}><summary><b>＋ Neues Turnier</b></summary>
    <form data-f="turnierNeu" autocomplete="off" style="margin-top:.6rem"><div class="grid2">
      <label>Name *<input name="name" required placeholder="z. B. Kids-Cup"></label><label>Datum<input type="date" name="d" value="${today()}" required></label>
      <label>Ort / Golfclub<input name="ort"></label><label>Löcher<select name="loecher"><option value="">–</option><option>6</option><option selected>9</option><option>18</option></select></label>
    </div><div class="knoepfe" style="margin-top:.8rem"><button class="pri" type="submit">Anlegen</button></div></form></details>
  ${L.length ? `<ul class="liste">${L.map(x => { const plaetze = x.tn.map(p => p.platz).filter(p => p != null); return `<li><a class="zeile" href="#/turnier/${h(x.id)}"><span><b>${esc(x.name || "Turnier")}</b>
    <small>${[fmtD(x.d), x.ort && esc(x.ort), x.loecher && x.loecher + " Löcher"].filter(Boolean).join(" · ")}</small></span>
    <span class="re"><b>${x.tn.length}</b> Kinder${plaetze.length ? `<br>bester Platz ${Math.min(...plaetze)}` : ""}</span></a></li>`; }).join("")}</ul>`
    : `<div class="leer"><p>Noch keine Turniere eingetragen.</p></div>`}
  ${L.length ? `<div class="karte" style="margin-top:1rem"><h3>Turniere je Kind</h3>${zeitraumWahl()}
    ${S.length ? `<ul class="namen" style="margin-top:.6rem">${S.map(s => `<li><a href="#/kind/${h(s.kind.id)}">${name(s.kind)}</a><span>${s.anzahl} Turnier${s.anzahl === 1 ? "" : "e"}${s.besterPlatz != null ? ` · bester Platz ${s.besterPlatz}` : ""}</span></li>`).join("")}</ul>` : `<p class="muted">Im Zeitraum keine Turniere.</p>`}</div>` : ""}`;
}
function seiteTurnier(id) {
  const x = turnier(id);
  if (!x) return nichtGefunden("Dieses Turnier gibt es nicht (mehr).", "#/turniere", "Alle Turniere");
  const tn = x.tn.slice().sort((a, b) => (a.platz ?? 999) - (b.platz ?? 999) || vergleich(M.kindName(M.kindVon(D, a.kind)), M.kindName(M.kindVon(D, b.kind))));
  const frei = M.kinderFiltern(D, { gruppe: A.gruppe }).filter(k => !x.tn.some(p => p.kind === k.id));
  const feld = (f, text, art = "text", extra = "") => `<label>${text}<input type="${art}" value="${esc(x[f] ?? "")}" data-c="wFeld" data-id="${esc(x.id)}" data-f="${f}" autocomplete="off"${extra}></label>`;
  const z = (p, f, text) => `<label>${text}<input type="text" inputmode="numeric" value="${p[f] != null ? esc(zahlText(p[f])) : ""}" data-c="tnWert" data-id="${esc(x.id)}" data-k="${esc(p.kind)}" data-f="${f}" autocomplete="off"></label>`;
  return `<a class="zurueck" href="#/turniere">‹ Alle Turniere</a>
  <div class="kopf"><h2>${esc(x.name || "Turnier")}</h2><span class="muted">${fmtDW(x.d)}</span></div>
  <div class="karte"><div class="grid2">${feld("name", "Name")}${feld("d", "Datum", "date")}${feld("ort", "Ort / Golfclub")}${feld("loecher", "Löcher", "text", ' inputmode="numeric"')}</div>
    <label style="margin-top:.7rem">Notiz<textarea data-i="wNotiz" data-id="${esc(x.id)}" rows="2">${esc(x.notiz)}</textarea></label></div>
  <h3 style="margin:0 0 .5rem">Ergebnisse unserer Kinder</h3>
  ${tn.length ? `<ul class="eingabe turnier">${tn.map(p => `<li><span class="wer"><b><a href="#/kind/${h(p.kind)}">${name(M.kindVon(D, p.kind))}</a></b></span>
    ${z(p, "schlaege", "Schläge")}${z(p, "punkte", "Punkte")}${z(p, "platz", "Platz")}
    <button class="klein leise" data-a="tnWeg" data-id="${esc(x.id)}" data-k="${esc(p.kind)}" aria-label="${name(M.kindVon(D, p.kind))} entfernen">✕</button></li>`).join("")}</ul>
    <p class="klein muted">Wird sofort gespeichert. Schläge und/oder Punkte (z. B. Stableford) – was es bei dem Turnier gab.</p>` : `<p class="muted">Noch keine Kinder eingetragen.</p>`}
  ${frei.length ? `<details class="karte" style="margin-top:1rem"${tn.length ? "" : " open"}><summary><b>Kinder hinzufügen</b></summary>
    <form data-f="tnDazu" data-id="${esc(x.id)}" style="margin-top:.5rem"><ul class="liste">${frei.map(k => `<li><label class="chk auswahl"><input type="checkbox" name="k" value="${esc(k.id)}"> ${name(k)}${gruppenKlein(k)}</label></li>`).join("")}</ul>
    <div class="knoepfe" style="margin-top:.8rem"><button class="pri" type="submit">Hinzufügen</button></div></form></details>` : ""}
  <div class="knoepfe" style="margin-top:1rem"><button class="gefahr" data-a="turnierLoeschen" data-id="${esc(x.id)}">Turnier löschen</button></div>`;
}

/* ----- Mehr ----- */
function seiteMehr() {
  const vorher = Sp.vorherLaden();
  return `<div class="kopf"><h2>Mehr</h2></div>
  ${googleBereich()}
  <section class="karte" id="sichern"><h3>Daten sichern (Excel)</h3>
    <p style="margin-top:0">${D.excelGesichert ? `Zuletzt als Excel gesichert: <b>${fmtDT(D.excelGesichert)}</b>.` : "Noch nie als Excel gesichert."}
      ${D.ungesichert && !mitGoogle() ? `<br><b class="warn">Es gibt Änderungen, die noch nicht als Excel gesichert sind.</b>` : ""}</p>
    <div class="knoepfe"><button class="pri" data-a="excelSpeichern">Als Excel speichern</button>
      <label class="knopf">Excel-Datei laden …<input type="file" accept=".xlsx,${XLSX_TYP}" data-c="excelDatei" class="versteckt"></label>
      ${ORDNER && ORDNER.stand ? `<button data-a="ordnerLaden">Neuesten Stand aus dem Ordner laden</button>` : ""}</div>
    <p class="klein muted">${ORDNER ? `Am PC wird jedes Mal eine neue Datei im Ordner <b>KISY › Daten</b> gespeichert – vorhandene Dateien werden nie überschrieben.${ORDNER.stand ? ` Neueste Datei: ${esc(ORDNER.stand)}.` : ""}`
      : IOS ? `Auf dem iPhone öffnet sich „Teilen“: z. B. „In Dateien sichern“ oder per AirDrop an den PC.` : `Die Datei landet im Download-Ordner.`}
      Excel laden ersetzt den Stand auf diesem Gerät (vorher wird gefragt).</p>
    ${vorher ? `<button class="leise klein" data-a="vorherHolen">Stand vor dem letzten Laden wiederherstellen</button>` : ""}
  </section>
  <section class="karte" id="listen"><h3>Listen bearbeiten</h3>
    <h4>Gruppen</h4>
    ${D.gruppen.length ? `<ul class="bearb">${D.gruppen.map(g => `<li><input type="text" value="${esc(g)}" data-c="gruppeName" data-alt="${esc(g)}" aria-label="Gruppe ${esc(g)}" autocomplete="off"><span class="klein muted">${M.gruppeAnzahl(D, g)} Kinder</span><button class="klein gefahr" data-a="gruppeWeg" data-g="${esc(g)}">Löschen</button></li>`).join("")}</ul>` : `<p class="muted klein">Keine Gruppen. Ohne Gruppen gehören alle Kinder zu jedem Training.</p>`}
    <p class="klein muted">Umbenennen: Namen ändern und das Feld verlassen. Welche Kinder in welchen Gruppen sind, stellst du beim Kind ein.</p>
    <form data-f="gruppeNeu" class="row" autocomplete="off"><label>Neue Gruppe<input name="name" required></label><button type="submit">Hinzufügen</button></form>

    <h4 style="margin-top:1.4rem">Abzeichen</h4>
    <ul class="bearb">${D.abzeichen.map(a => `<li><input type="text" value="${esc(a.name)}" data-c="aFeld" data-id="${esc(a.id)}" data-f="name" aria-label="Abzeichen" autocomplete="off">
      <label class="chk klein"><input type="checkbox" data-c="aFeld" data-id="${esc(a.id)}" data-f="aus"${a.aus ? " checked" : ""}> ausgeblendet</label>
      ${M.abzeichenBenutzt(D, a.id) ? "" : `<button class="klein gefahr" data-a="abzWeg" data-id="${esc(a.id)}">Löschen</button>`}</li>`).join("")}</ul>
    <form data-f="abzNeu" class="row" autocomplete="off"><label>Neues Abzeichen<input name="name" required></label><button type="submit">Hinzufügen</button></form>

    <h4 style="margin-top:1.4rem">Einverständnisse</h4>
    <ul class="bearb">${D.einvArten.map(e => `<li><input type="text" value="${esc(e.name)}" data-c="eFeld" data-id="${esc(e.id)}" aria-label="Einverständnis" autocomplete="off"><button class="klein gefahr" data-a="einvWeg" data-id="${esc(e.id)}">Löschen</button></li>`).join("")}</ul>
    <form data-f="einvNeu" class="row" autocomplete="off"><label>Neue Art<input name="name" required placeholder="z. B. Mitgliedsantrag"></label><button type="submit">Hinzufügen</button></form>
  </section>
  <section class="karte"><h3>Datenschutz</h3>
    <p style="margin-top:0">Alle Angaben bleiben auf diesem Gerät, in deinen Excel-Dateien und – wenn verbunden – in deinem eigenen Google Drive. Sonst schickt KISY nichts ins Internet.</p>
    <button class="gefahr" data-a="allesLoeschen">Alle KISY-Daten auf diesem Gerät löschen</button></section>
  <p class="klein muted">KISY Version ${VERSION}</p>`;
}

function googleBereich() {
  if (!G.konfiguriert()) return "";
  if (!G.warVerbunden()) return `<section class="karte" id="google"><h3>Google Drive</h3>
    <p style="margin-top:0">Verbinde KISY mit deinem Google-Konto, dann haben iPhone und PC immer denselben Stand. Gespeichert wird in deinem Google Drive im Ordner <b>${G.ORDNER}</b> (Datei „${G.STAND_DATEI}“), dazu einmal pro Woche eine Sicherung.</p>
    <button class="pri" data-a="gVerbinden">Mit Google verbinden</button>
    <p class="klein muted">KISY bekommt nur Zugriff auf die Dateien, die es selbst anlegt – nicht auf deine anderen Dateien. Der Zugang bleibt auf diesem Gerät.</p></section>`;
  return `<section class="karte" id="google"><h3>Google Drive</h3>
    <p style="margin-top:0">${G.angemeldet() ? `Verbunden${G.email() ? " als <b>" + esc(G.email()) + "</b>" : ""}.` : `<b class="warn">Die Google-Anmeldung ist abgelaufen.</b> Deine Eingaben bleiben auf dem Gerät und werden nach dem Verbinden übertragen.`}
      ${SYNC.zuletzt ? `<br>Zuletzt abgeglichen: ${fmtDT(SYNC.zuletzt)}.` : ""}</p>
    <div class="knoepfe">${G.angemeldet() ? `<button class="pri" data-a="jetztAbgleichen">Jetzt abgleichen</button><button data-a="jetztSichern">Jetzt sichern</button>` : `<button class="pri" data-a="gVerbinden">Neu verbinden</button>`}
      <button class="leise" data-a="gTrennen">Verbindung trennen</button></div>
    <p class="klein muted">Stand: Google Drive › ${G.ORDNER} › ${G.STAND_DATEI} (am PC auch mit Excel zu öffnen). Sicherungen: einmal pro Woche in ${G.ORDNER} › ${G.SICHERUNG_ORDNER}${D.sicherung ? `, zuletzt am ${fmtD(D.sicherung)}` : ""}.
      Änderungen von iPhone und PC werden automatisch zusammengeführt.</p></section>`;
}

function trainingOeffnen(gruppe) {
  const x = M.trainingFuer(D, today(), gruppe);
  A.heute = false; persist(); geh("#/training/" + encodeURIComponent(x.id));
}

/* ---------- Klicks ---------- */
const KLICK = {
  zuMehr() { if (mitGoogle() && !G.angemeldet() && navigator.onLine) { jetztSpeichern(); G.anmelden(false); } else geh("#/mehr"); },
  gVerbinden() { jetztSpeichern(); G.anmelden(false); },
  gTrennen() {
    if (!confirm("Verbindung zu Google trennen?\n\nDie Daten bleiben auf diesem Gerät und in Google Drive. Bis zum erneuten Verbinden wird nicht mehr abgeglichen.")) return;
    G.abmelden(); toast("Verbindung zu Google getrennt"); render();
  },
  jetztAbgleichen() { abgleichen(true); },
  async jetztSichern() {
    try { await sicherungAnlegen(`KISY_Sicherung_${zeitstempel()}.xlsx`); toast("Sicherung in Google Drive angelegt"); render(); }
    catch (e) { toast("Sicherung nicht möglich: " + e.message, 5000); }
  },
  trainingHeute() {
    if (!D.kinder.length && !confirm("Es sind noch keine Kinder angelegt. Trotzdem ein Training anlegen?")) return;
    if (D.gruppen.length) { A.heute = !A.heute; if (route().teil === "training" && !route().rest.length) render(); else geh("#/training"); window.scrollTo(0, 0); return; }
    trainingOeffnen("");
  },
  heuteGruppe(el) { trainingOeffnen(el.dataset.g); },
  heuteZu() { A.heute = false; render(); },
  anw(el) {
    const x = training(el.dataset.id); if (!x) return;
    const da = el.getAttribute("aria-pressed") !== "true";
    M.anwesenheitSetzen(x, el.dataset.k, da); persist();
    document.querySelectorAll(`.anw[data-k="${CSS.escape(el.dataset.k)}"]`).forEach(b => b.setAttribute("aria-pressed", String(da)));
    $("#zaehler").innerHTML = zaehlerText(x);
  },
  alleDa(el) { const x = training(el.dataset.id); M.kinderFuerTraining(D, x).filter(k => k.aktiv).forEach(k => M.anwesenheitSetzen(x, k.id, true)); persist(); render(); },
  keinerDa(el) { const x = training(el.dataset.id); if (x.da.length > 3 && !confirm("Alle Häkchen dieses Trainings entfernen?")) return; x.da = []; persist(); render(); },
  trainingLoeschen(el) {
    const x = training(el.dataset.id);
    if (!confirm(`Training vom ${fmtD(x.d)} löschen?${x.da.length ? ` Die Anwesenheit (${x.da.length} Kinder) geht dabei verloren.` : ""}`)) return;
    D.trainings = D.trainings.filter(t => t !== x); persist(); toast("Training gelöscht"); geh("#/training");
  },
  kindLoeschen(el) {
    const k = M.kindVon(D, el.dataset.id);
    if (!confirm(`${M.kindName(k)} wirklich löschen?\n\nDabei werden auch Anwesenheit, Abzeichen und Turniere dieses Kindes gelöscht.\n\nTipp: Wenn das Kind nur nicht mehr kommt, stattdessen „aktiv“ ausschalten.`)) return;
    M.kindLoeschen(D, k.id); persist(); toast("Gelöscht"); geh("#/kinder");
  },
  hcpWeg(el) { const k = M.kindVon(D, el.dataset.id); if (!confirm(`Handicap vom ${fmtD(el.dataset.d)} löschen?`)) return; k.hcp = k.hcp.filter(x => x.d !== el.dataset.d); persist(); render(); },
  erreichtWeg(el) { if (!confirm("Dieses Abzeichen beim Kind löschen?")) return; D.erreicht = D.erreicht.filter(e => e.id !== el.dataset.id); persist(); render(); },
  tnWeg(el) { const x = turnier(el.dataset.id); if (!confirm(`${M.kindName(M.kindVon(D, el.dataset.k))} aus diesem Turnier entfernen?`)) return; x.tn = x.tn.filter(p => p.kind !== el.dataset.k); persist(); render(); },
  turnierLoeschen(el) { const x = turnier(el.dataset.id); if (!confirm(`Turnier „${x.name}“ mit allen Ergebnissen löschen?`)) return; D.turniere = D.turniere.filter(t => t !== x); persist(); toast("Turnier gelöscht"); geh("#/turniere"); },
  gruppeWeg(el) {
    const g = el.dataset.g, n = M.gruppeAnzahl(D, g);
    if (!confirm(`Gruppe „${g}“ löschen?${n ? ` Sie wird bei ${n} Kindern entfernt.` : ""} Trainings dieser Gruppe gelten danach für alle Kinder.`)) return;
    M.gruppeLoeschen(D, g); if (A.gruppe === g) A.gruppe = ""; persist(); render();
  },
  abzWeg(el) { if (!confirm("Dieses Abzeichen aus der Liste löschen?")) return; D.abzeichen = D.abzeichen.filter(a => a.id !== el.dataset.id); persist(); render(); },
  einvWeg(el) {
    const e = D.einvArten.find(x => x.id === el.dataset.id), n = D.kinder.filter(k => k.einv[e.id]).length;
    if (!confirm(`„${e.name}“ löschen?${n ? ` Die Häkchen bei ${n} Kindern gehen dabei verloren.` : ""}`)) return;
    D.einvArten = D.einvArten.filter(x => x !== e); for (const k of D.kinder) delete k.einv[e.id]; persist(); render();
  },
  excelSpeichern() { excelSpeichern(); },
  async ordnerLaden() {
    try { const r = await Sp.ausOrdnerLaden(); await excelEinlesen(r.daten, r.name); }
    catch (e) { toast("Laden nicht möglich: " + e.message, 5000); }
  },
  vorherHolen() {
    const v = Sp.vorherLaden(); if (!v) return;
    if (!confirm(`Den Stand von vor dem letzten Excel-Laden wiederherstellen?\n\nDarin: ${M.zaehlenText(M.zaehlen(M.normalisieren(v)))}\nJetzt: ${M.zaehlenText(M.zaehlen(D))}`)) return;
    D = { ...M.normalisieren(v), drive: D.drive, ersetzen: mitGoogle() }; persist(); toast("Vorheriger Stand ist wieder da"); render();
  },
  allesLoeschen() {
    if (!confirm(`Wirklich ALLE KISY-Daten auf diesem Gerät löschen?\n\n${mitGoogle() ? "Der Stand in Google Drive bleibt erhalten; die Verbindung zu Google wird auf diesem Gerät getrennt." : "Hast du vorher als Excel gesichert? Ohne Sicherung sind die Daten weg."}`)) return;
    if (prompt("Zur Sicherheit bitte LÖSCHEN eintippen:") !== "LÖSCHEN") { toast("Nicht gelöscht"); return; }
    clearTimeout(syncTimer); if (G.warVerbunden()) G.abmelden();
    Sp.allesLoeschen(); D = M.neueDaten(); persist(false); toast("Alle Daten auf diesem Gerät gelöscht"); geh("#/training");
  },
};

/* ---------- Änderungen in Feldern ---------- */
const AENDERUNG = {
  zr(el) { A.zr = el.value; render(); },
  zrVon(el) { A.von = el.value; render(); },
  zrBis(el) { A.bis = el.value; render(); },
  gruppe(el) { A.gruppe = el.value; render(); },
  sort(el) { A.sort = el.value; render(); },
  inaktive(el) { A.inaktive = el.checked; render(); },
  tDatum(el) { const x = training(el.dataset.id); if (!istDatum(el.value)) { el.value = x.d; return; } x.d = el.value; persist(); render(); },
  tGruppe(el) { const x = training(el.dataset.id); x.gruppe = el.value; persist(); render(); },
  wFeld(el) {
    const x = turnier(el.dataset.id), f = el.dataset.f;
    if (f === "d") { if (!istDatum(el.value)) { el.value = x.d; return; } x.d = el.value; }
    else if (f === "loecher") x.loecher = el.value.trim() ? M.zahlPruefen(el.value).wert ?? x.loecher : null;
    else x[f] = el.value.trim();
    persist(); if (f === "d" || f === "name") render();
  },
  tnWert(el) {
    const x = turnier(el.dataset.id), p = x.tn.find(q => q.kind === el.dataset.k), r = M.zahlPruefen(el.value);
    if (r.fehler) { el.classList.add("falsch"); toast(r.fehler); return; }
    el.classList.remove("falsch"); p[el.dataset.f] = r.wert; persist();
  },
  einvHaken(el) { const d = el.form.elements["einvd_" + el.name.slice(5)]; if (el.checked && d && !d.value) d.value = today(); if (!el.checked && d) d.value = ""; },
  async excelDatei(el) {
    const f = el.files && el.files[0]; el.value = ""; if (!f) return;
    await excelEinlesen(await f.arrayBuffer(), f.name);
  },
  gruppeName(el) {
    const neu = el.value.replace(/,/g, " ").trim(), alt = el.dataset.alt;
    if (!neu) { el.value = alt; return; }
    if (D.gruppen.includes(neu) && neu !== alt && !confirm(`Die Gruppe „${neu}“ gibt es schon. Beide zusammenlegen?`)) { el.value = alt; return; }
    M.gruppeUmbenennen(D, alt, neu); if (A.gruppe === alt) A.gruppe = neu; persist(); render();
  },
  aFeld(el) {
    const a = M.abzeichenVon(D, el.dataset.id);
    if (el.dataset.f === "aus") a.aus = el.checked; else { if (!el.value.trim()) { el.value = a.name; return; } a.name = el.value.trim(); }
    persist(); toast("Gespeichert", 1200);
  },
  eFeld(el) { const e = D.einvArten.find(x => x.id === el.dataset.id); if (!el.value.trim()) { el.value = e.name; return; } e.name = el.value.trim(); persist(); toast("Gespeichert", 1200); },
};

/* ---------- Eingaben (während des Tippens) ---------- */
// Notizen werden kurz nach dem Tippen gespeichert – spätestens beim Seitenwechsel oder wenn die App in den Hintergrund geht.
let tippTimer = null, tippFn = null;
const spaeter = fn => { clearTimeout(tippTimer); tippFn = fn; tippTimer = setTimeout(jetztSpeichern, 400); };
function jetztSpeichern() { clearTimeout(tippTimer); const fn = tippFn; tippFn = null; if (fn) fn(); }
const EINGABE = {
  q(el) { A.q = el.value; $("#kliste").innerHTML = kinderZeilen(); },
  tNotiz(el) { const id = el.dataset.id; spaeter(() => { const x = training(id); if (x) { x.notiz = el.value; persist(); } }); },
  wNotiz(el) { const id = el.dataset.id; spaeter(() => { const x = turnier(id); if (x) { x.notiz = el.value.trim(); persist(); } }); },
};

/* ---------- Formulare ---------- */
const FORMULAR = {
  trainingNeu(f) {
    const d = f.elements.d.value; if (!istDatum(d)) return;
    const x = M.neuesTraining(D, d, f.elements.gruppe ? f.elements.gruppe.value : "");
    persist(); geh("#/training/" + encodeURIComponent(x.id));
  },
  kind(f) {
    const w = Object.fromEntries(new FormData(f)), v = (w.vorname || "").trim();
    if (!v) { toast("Bitte einen Vornamen eingeben"); return; }
    let k = f.dataset.id ? M.kindVon(D, f.dataset.id) : null;
    const neu = !k;
    if (neu) { k = M.neuesKind(); D.kinder.push(k); }
    for (const feld of ["vorname", "nachname", "geburt", "seit", "notiz", "eltern", "elternTel", "elternMail", "notfall", "notfallTel", "hinweise"]) k[feld] = (w[feld] || "").trim();
    k.gruppen = new FormData(f).getAll("gruppe");
    k.aktiv = !!w.aktiv;
    k.einv = {};
    for (const e of D.einvArten) if (w["einv_" + e.id]) k.einv[e.id] = istDatum(w["einvd_" + e.id]) ? w["einvd_" + e.id] : "ja";
    persist(); toast(neu ? `${M.kindName(k)} angelegt` : "Gespeichert"); geh("#/kind/" + encodeURIComponent(k.id));
  },
  mehrere(f) {
    const L = M.namenLesen(f.elements.namen.value), gruppen = new FormData(f).getAll("gruppe"), seit = f.elements.seit.value;
    if (!L.length) return;
    const doppelt = L.filter(n => D.kinder.some(k => k.vorname.toLowerCase() === n.vorname.toLowerCase() && k.nachname.toLowerCase() === n.nachname.toLowerCase()));
    if (doppelt.length && !confirm(`Schon vorhanden: ${doppelt.map(n => n.vorname + " " + n.nachname).join(", ")}.\nTrotzdem alle anlegen?`)) return;
    for (const n of L) D.kinder.push(M.neuesKind({ ...n, gruppen, seit: istDatum(seit) ? seit : today() }));
    persist(); toast(`${L.length} Kinder angelegt`); geh("#/kinder");
  },
  hcp(f) {
    const k = M.kindVon(D, f.dataset.id), d = f.elements.d.value, w = M.hcpLesen(f.elements.w.value);
    if (!istDatum(d)) return;
    if (w == null || w > 54 || w < -10) { toast("Bitte ein Handicap zwischen +10 und 54 eingeben"); return; }
    k.hcp = k.hcp.filter(x => x.d !== d).concat({ d, w }).sort((a, b) => a.d.localeCompare(b.d));
    persist(); toast("Handicap eingetragen"); render();
  },
  kindAbz(f) {
    const d = f.elements.d.value; if (!istDatum(d)) return;
    M.abzeichenEintragen(D, f.elements.abz.value, d, [f.dataset.id]); persist(); toast("Abzeichen eingetragen"); render();
  },
  abzEintragen(f) {
    const d = f.elements.d.value, kinder = new FormData(f).getAll("k");
    if (!istDatum(d)) return;
    if (!kinder.length) { toast("Bitte mindestens ein Kind ankreuzen"); return; }
    const n = M.abzeichenEintragen(D, f.dataset.id, d, kinder); persist(); toast(`${n} eingetragen`); geh("#/abzeichen");
  },
  turnierNeu(f) {
    const name = f.elements.name.value.trim(), d = f.elements.d.value; if (!name || !istDatum(d)) return;
    const x = M.neuesTurnier(D, { name, d, ort: f.elements.ort.value, loecher: f.elements.loecher.value ? +f.elements.loecher.value : null });
    persist(); geh("#/turnier/" + encodeURIComponent(x.id));
  },
  tnDazu(f) {
    const x = turnier(f.dataset.id), kinder = new FormData(f).getAll("k");
    if (!kinder.length) { toast("Bitte mindestens ein Kind ankreuzen"); return; }
    for (const kind of kinder) if (!x.tn.some(p => p.kind === kind)) x.tn.push({ kind, schlaege: null, punkte: null, platz: null, notiz: "" });
    persist(); render();
  },
  gruppeNeu(f) { const g = f.elements.name.value.replace(/,/g, " ").trim(); if (!g) return; if (!D.gruppen.includes(g)) D.gruppen = D.gruppen.concat(g); persist(); render(); },
  abzNeu(f) { const name = f.elements.name.value.trim(); if (!name) return; D.abzeichen.push({ id: "a-" + Date.now().toString(36), name, aus: false }); persist(); render(); },
  einvNeu(f) { const name = f.elements.name.value.trim(); if (!name) return; D.einvArten.push({ id: "e-" + Date.now().toString(36), name }); persist(); render(); },
};

/* ---------- Excel ---------- */
async function excelSpeichern() {
  const daten = schreibeStand(XLSX, D), dateiname = `KISY_Stand_${zeitstempel()}.xlsx`;
  try {
    if (ORDNER) {
      const datei = await Sp.inOrdnerSpeichern(daten);
      gesichert(); toast("Gespeichert: KISY › Daten › " + datei, 4000);
      ORDNER = await Sp.ordnerVerfuegbar() || ORDNER; render(); return;
    }
    const datei = new File([daten], dateiname, { type: XLSX_TYP });
    if (IOS && navigator.canShare && navigator.canShare({ files: [datei] })) await navigator.share({ files: [datei], title: dateiname });
    else {
      const a = document.createElement("a"); a.href = URL.createObjectURL(datei); a.download = dateiname;
      document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
    }
    gesichert(); toast("Excel-Datei erstellt: " + dateiname, 4000); render();
  } catch (e) {
    if (e && e.name === "AbortError") return; // Teilen abgebrochen
    toast("Speichern nicht möglich: " + (e.message || e), 5000);
  }
}
function gesichert() { D.ungesichert = false; D.excelGesichert = new Date().toISOString(); persist(false); }
async function excelEinlesen(puffer, dateiname) {
  let r;
  try { r = leseStand(XLSX, puffer); } catch (e) { toast(e.message || "Die Datei konnte nicht gelesen werden", 5000); return; }
  if (!confirm(`Excel-Stand „${dateiname}“ laden?\n\nIn der Datei: ${M.zaehlenText(M.zaehlen(r.daten))}\nJetzt auf diesem Gerät: ${M.zaehlenText(M.zaehlen(D))}\n\nDer Stand auf diesem Gerät wird dabei ersetzt. Er wird vorher aufgehoben und kann unter „Mehr“ wiederhergestellt werden.${mitGoogle() ? "\n\nAuch der Stand in Google Drive wird ersetzt (dort wird vorher eine Sicherung angelegt)." : ""}`)) return;
  Sp.sichern();
  const drive = D.drive;
  D = r.daten; D.ungesichert = false; D.excelGesichert = new Date().toISOString();
  if (mitGoogle()) { D.drive = drive; D.ersetzen = true; D.driveOffen = true; planeAbgleich(500); }
  persist(false);
  toast("Excel-Stand geladen"); render();
  if (r.hinweise.length) alert("Hinweise beim Laden:\n\n" + r.hinweise.slice(0, 20).join("\n") + (r.hinweise.length > 20 ? `\n… und ${r.hinweise.length - 20} weitere` : ""));
}

/* ---------- Abgleich mit Google Drive ---------- */
const DRIVE = {
  info: () => G.standInfo(), laden: id => G.standLaden(id), speichern: (inhalt, id) => G.standSpeichern(inhalt, id),
  sichern: (inhalt, zusatz) => G.sicherungSpeichern(`KISY_Sicherung_${zeitstempel()}_${zusatz}.xlsx`, inhalt),
};
function planeAbgleich(ms = 3000) { if (!G.angemeldet()) return; clearTimeout(syncTimer); syncTimer = setTimeout(abgleichen, ms); }
async function abgleichen(manuell = false) {
  if (!G.konfiguriert()) return;
  if (!G.angemeldet()) { statusZeigen(); if (manuell === true) G.anmelden(false); return; }
  if (SYNC.laeuft) return;
  clearTimeout(syncTimer); jetztSpeichern();
  SYNC.laeuft = true; SYNC.fehler = ""; statusZeigen();
  try {
    const stand = D.aenderung || 0, vorher = structuredClone(D);
    const r = await AG.abgleich({ daten: vorher, basis: Sp.basisLaden() }, DRIVE, XLSX);
    if ((D.aenderung || 0) === stand) D = { ...r.daten, aenderung: stand };
    else { // während des Abgleichs wurde weiter eingegeben: diese Änderungen kommen obendrauf
      const neu = AG.zusammenfuehren(AG.nurInhalt(vorher), D, r.daten);
      D = { ...neu, ungesichert: D.ungesichert, excelGesichert: D.excelGesichert, geaendert: D.geaendert, aenderung: D.aenderung, drive: r.daten.drive, driveOffen: true };
    }
    Sp.basisSpeichern(r.basis); SYNC.zuletzt = new Date().toISOString();
    persist(false);
    if (r.aktion === "geladen" || r.aktion === "zusammengefuehrt") { neuZeichnen(); toast("Neuer Stand aus Google Drive übernommen"); }
    else if (manuell === true) toast(r.aktion === "nichts" ? "Alles ist auf dem neuesten Stand" : "In Google Drive gespeichert");
    await wochenSicherung();
  } catch (e) {
    console.warn(e);
    SYNC.fehler = e.code === "anmelden" ? "" : navigator.onLine ? "Google Drive nicht erreichbar – wird erneut versucht" : "";
    if (manuell === true && e.code !== "anmelden") toast(SYNC.fehler || e.message, 5000);
  } finally {
    SYNC.laeuft = false; statusZeigen();
    if (route().teil === "mehr") neuZeichnen();
    if (SYNC.fehler) planeAbgleich(60e3);
    else if (D.driveOffen) planeAbgleich();
  }
}
// Sicherungskopie des ganzen Stands in Google Drive › KISY › Sicherungen
async function sicherungAnlegen(dateiname) {
  await G.sicherungSpeichern(dateiname, schreibeStand(XLSX, D));
  D.sicherung = today(); persist(); // geht mit in den Stand, damit nicht jedes Gerät noch einmal sichert
}
// Einmal pro Woche (ab Montag) beim Abgleich
async function wochenSicherung() {
  if (D.sicherung && D.sicherung >= montag(today())) return;
  try { await sicherungAnlegen(`KISY_Sicherung_${today()}.xlsx`); } catch (e) { console.warn("Sicherung", e); }
}
// Nicht neu aufbauen, solange gerade in ein Feld geschrieben wird – dann erst danach
function neuZeichnen() {
  const a = document.activeElement;
  if (a && a.closest && a.closest("main") && a.matches("input, textarea, select")) { NEU_ZEICHNEN = true; return; }
  NEU_ZEICHNEN = false; render();
}
// Ist die Google-Anmeldung abgelaufen, still neu anmelden (höchstens alle 10 Minuten, nicht mitten in einem Formular)
function stillAnmelden() {
  if (!mitGoogle() || G.angemeldet() || !navigator.onLine) return false;
  const r = route();
  if ((r.teil === "kind" && (r.rest[0] === "neu" || r.rest[1] === "bearbeiten")) || (r.teil === "kinder" && r.rest[0] === "mehrere")) return false;
  let zuletzt = 0; try { zuletzt = +sessionStorage.getItem("kisy-still") || 0; } catch (e) { /* egal */ }
  if (Date.now() - zuletzt < 10 * 60e3) return false;
  try { sessionStorage.setItem("kisy-still", String(Date.now())); } catch (e) { /* egal */ }
  jetztSpeichern(); G.anmelden(true); return true;
}

/* ---------- Start ---------- */
function verbinden() {
  document.body.addEventListener("click", e => {
    const el = e.target.closest("[data-a]"); if (!el || !KLICK[el.dataset.a]) return;
    e.preventDefault(); KLICK[el.dataset.a](el, e);
  });
  document.body.addEventListener("change", e => { const el = e.target.closest("[data-c]"); if (el && AENDERUNG[el.dataset.c]) AENDERUNG[el.dataset.c](el, e); });
  document.body.addEventListener("input", e => { const el = e.target.closest("[data-i]"); if (el && EINGABE[el.dataset.i]) EINGABE[el.dataset.i](el, e); });
  document.body.addEventListener("submit", e => { const f = e.target.closest("form[data-f]"); if (!f) return; e.preventDefault(); if (FORMULAR[f.dataset.f]) FORMULAR[f.dataset.f](f, e); });
  // Turnier-Ergebnisse: Enter springt zum nächsten Feld
  document.body.addEventListener("keydown", e => {
    if (e.key !== "Enter" || !e.target.matches(".eingabe input")) return;
    e.preventDefault(); const alle = [...document.querySelectorAll(".eingabe input")], i = alle.indexOf(e.target);
    if (alle[i + 1]) { alle[i + 1].focus(); alle[i + 1].scrollIntoView({ block: "center", behavior: "smooth" }); } else e.target.blur();
  });
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) { jetztSpeichern(); return; }
    if (!stillAnmelden()) abgleichen();
  });
  document.body.addEventListener("focusout", () => setTimeout(() => { if (NEU_ZEICHNEN) neuZeichnen(); }, 300));
  addEventListener("online", () => { statusZeigen(); abgleichen(); });
  addEventListener("offline", statusZeigen);
  window.addEventListener("pagehide", jetztSpeichern);
  window.addEventListener("hashchange", () => {
    jetztSpeichern();
    render(); window.scrollTo(0, 0);
    if (location.hash === "#/mehr/listen") $("#listen")?.scrollIntoView();
  });
  // In einem zweiten Fenster geändert? Dann den neuen Stand übernehmen.
  window.addEventListener("storage", e => { if (e.key === "kisy-v1" && e.newValue) { D = M.normalisieren(JSON.parse(e.newValue)); render(); } });
}
async function start() {
  const rueck = G.rueckkehrVerarbeiten(); // zurück von der Google-Anmeldung?
  const s = Sp.laden();
  D = s ? M.normalisieren(s) : M.neueDaten();
  if (!s || (s.version || 1) < M.FORMAT) SPEICHER_OK = Sp.speichern(D); // neu oder gerade auf das neue Format umgestellt
  verbinden(); render();
  if (location.hash === "#/mehr/listen") $("#listen")?.scrollIntoView();
  ORDNER = await Sp.ordnerVerfuegbar();
  if (ORDNER) render();
  Sp.dauerhaftBitten();
  if (rueck && rueck.fehler && !/interaction_required|login_required|consent_required/.test(rueck.fehler)) toast("Google-Anmeldung nicht erfolgt (" + rueck.fehler + ")", 5000);
  if (G.konfiguriert()) {
    if (G.angemeldet()) {
      if (rueck && rueck.ok) { await G.kontoLaden().catch(() => {}); toast("Mit Google verbunden – Stand wird abgeglichen"); }
      abgleichen();
    } else if (!rueck) stillAnmelden();
    setInterval(() => { if (document.visibilityState === "visible") abgleichen(); }, 2 * 60e3);
  }
  if ("serviceWorker" in navigator && (location.protocol === "https:" || /^(localhost|127\.0\.0\.1)$/.test(location.hostname)))
    navigator.serviceWorker.register("sw.js").catch(() => { /* ohne Offline-Speicher geht es auch */ });
}
start();
