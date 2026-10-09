// Kleine Verlaufslinie (SVG) für Übungsergebnisse und Handicap.
// werte: [{ d, w }] nach Datum sortiert. Gleichmäßiger Abstand zwischen den Punkten.
import { esc, fmtD, zahlText } from "./grundlagen.js";

export function verlauf(werte, { breite = 120, hoehe = 32, min = null, max = null, titel = "" } = {}) {
  if (!werte || !werte.length) return "";
  const W = werte.map(x => x.w);
  let lo = min ?? Math.min(...W), hi = max ?? Math.max(...W);
  if (hi === lo) { hi += 1; lo -= 1; }
  const r = 3, x = i => werte.length === 1 ? breite / 2 : r + i * (breite - 2 * r) / (werte.length - 1);
  const y = w => r + (hi - Math.min(hi, Math.max(lo, w))) * (hoehe - 2 * r) / (hi - lo);
  const pts = werte.map((p, i) => `${x(i).toFixed(1)},${y(p.w).toFixed(1)}`);
  const beschr = titel || werte.map(p => `${fmtD(p.d)}: ${zahlText(p.w)}`).join(", ");
  return `<svg class="verlauf" viewBox="0 0 ${breite} ${hoehe}" width="${breite}" height="${hoehe}" role="img" aria-label="${esc(beschr)}">`
    + (werte.length > 1 ? `<polyline points="${pts.join(" ")}" />` : "")
    + pts.map((p, i) => `<circle cx="${p.split(",")[0]}" cy="${p.split(",")[1]}" r="${i === pts.length - 1 ? 3 : 2}"${i === pts.length - 1 ? ' class="letzter"' : ""}><title>${esc(fmtD(werte[i].d) + ": " + zahlText(werte[i].w))}</title></circle>`).join("")
    + `</svg>`;
}
