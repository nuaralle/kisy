// Kleine Hilfen für Datum, Zahlen und Text. Datumswerte sind im Programm immer Text "JJJJ-MM-TT".
const z2 = n => String(n).padStart(2, "0");
export const iso = d => `${d.getFullYear()}-${z2(d.getMonth() + 1)}-${z2(d.getDate())}`;
export const parseISO = s => { const [y, m, d] = String(s).split("-").map(Number); return new Date(y, m - 1, d); };
export const today = () => iso(new Date());
export const addDays = (s, n) => { const d = parseISO(s); d.setDate(d.getDate() + n); return iso(d); };
export const addMonths = (s, n) => { const d = parseISO(s); const tag = d.getDate(); d.setDate(1); d.setMonth(d.getMonth() + n); d.setDate(Math.min(tag, new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate())); return iso(d); };
export const istDatum = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ""));
// "2026-10-09" -> "09.10.2026"
export const fmtD = s => istDatum(s) ? `${s.slice(8, 10)}.${s.slice(5, 7)}.${s.slice(0, 4)}` : "–";
// "2026-10-09" -> "Fr, 09.10.2026"
const WT = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
export const fmtDW = s => istDatum(s) ? `${WT[parseISO(s).getDay()]}, ${fmtD(s)}` : "–";
// Zeitstempel für Dateinamen: "2026-10-09_1734"
export function zeitstempel(d = new Date()) { return `${iso(d)}_${z2(d.getHours())}${z2(d.getMinutes())}`; }

// Zahl aus Eingabe ("36,4", "36.4", " 7 ") -> Zahl oder null
export function zahl(v) {
  if (v == null) return null;
  if (typeof v === "number") return isFinite(v) ? v : null;
  const s = String(v).trim().replace(/\s/g, "").replace(",", ".");
  if (!s || !/^[+-]?\d*\.?\d+$/.test(s)) return null;
  return parseFloat(s);
}
// Zahl -> deutsche Schreibweise ("36,4"), ganze Zahlen ohne Komma
export const zahlText = n => n == null || !isFinite(n) ? "" : (Math.round(n * 10) / 10).toLocaleString("de-DE", { maximumFractionDigits: 1 });

// Text für HTML sicher machen
export const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
// Kurze eindeutige Kennung, z. B. "k-lq2m7x4a"
export const neueId = (vorsilbe = "x") => `${vorsilbe}-${Date.now().toString(36).slice(-5)}${Math.random().toString(36).slice(2, 7)}`;
// Sortieren nach deutschem Alphabet (Ä wie A usw.)
export const vergleich = (a, b) => String(a || "").localeCompare(String(b || ""), "de", { sensitivity: "base" });
