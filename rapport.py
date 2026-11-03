"""Laver rapport.html: indkøb set i forhold til budgettet for de gæster, der har spist (pr. måned og pr. dag).

Tallene læses fra månedsfanerne i Kokkenes k ark.xlsx – samme beregning som arket:
  Budget  = Kuverter × kuvertpris + vareforbrug fra ekstrasalg   (kolonne Q + AD)
  Indkøb  = Indkøb i alt                                           (kolonne I)
  Kr. pr. kuvert = (indkøb − til vareforbrug) / kuverter           (kolonne R)
Siden laves igen ved hver timekørsel af ordre_til_ark.py.

  python rapport.py          # lav rapport.html nu
"""
import json
from datetime import datetime
from pathlib import Path

ROD = Path(__file__).resolve().parent
HTML = ROD / "rapport.html"
MDR = ["Januar", "Februar", "Marts", "April", "Maj", "Juni", "Juli", "August",
       "September", "Oktober", "November", "December"]
TOLERANCE = 0.05          # inden for ±5 % af budgettet tæller som "på budget"

# Kolonner i månedsfanerne (1 = A)
A, I, P, Q, AD = 1, 9, 16, 17, 30
KAT = range(2, 9)          # B..H: leverandørkolonnerne, der tilsammen er "Indkøb i alt"


def tal(v):
    """Excel-celle -> float. Valutaformaterede celler kommer som Decimal via COM; tekst og tomme tæller som 0."""
    if isinstance(v, str) or v is None:
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def hent_data(wb):
    """Læser dagstal fra alle månedsfaner i en åben Excel-projektmappe (COM)."""
    maaneder = []
    for nr, md in enumerate(MDR, 1):
        ws = wb.Worksheets(md)
        raekker = ws.Range("A9:AE39").Value
        dage = []
        for r in raekker:
            if r[A - 1] in (None, ""):
                continue
            dage.append({"dag": int(r[A - 1]), "indkoeb": round(tal(r[I - 1]), 2), "kuverter": int(tal(r[P - 1])),
                         "budget": round(tal(r[Q - 1]) + tal(r[AD - 1]), 2), "vareforbrug": round(tal(r[AD - 1]), 2),
                         "kat": [round(tal(r[c - 1]), 2) for c in KAT]})
        maaneder.append({"nr": nr, "navn": md, "dage": dage})
    kuvertpris = tal(wb.Worksheets(MDR[0]).Range("C4").Value)
    kantine = wb.Worksheets(MDR[0]).Range("B3").Value or ""
    kolonner = [str(wb.Worksheets(MDR[0]).Cells(8, c).Value or "") for c in KAT]
    return {"kuvertpris": kuvertpris, "kantine": kantine, "tolerance": TOLERANCE, "kolonner": kolonner,
            "aar": datetime.now().year, "opdateret": datetime.now().strftime("%d-%m-%Y kl. %H:%M"),
            "maaneder": maaneder}


def skriv_html(data, sti=HTML):
    html = SKABELON.replace("/*DATA*/null", json.dumps(data, ensure_ascii=False))
    sti.write_text(html, encoding="utf-8")
    return sti


SKABELON = r"""<!doctype html>
<html lang="da">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Køkkenets indkøb</title>
<style>
:root {
  color-scheme: light;
  --bg: #f6f6f4; --surface: #fcfcfb; --border: #e4e3df;
  --ink: #0b0b0b; --ink-2: #52514e; --ink-3: #85847f; --grid: #ecebe7;
  --over: #e34948; --under: #2a78d6; --neutral: #b9b8b2;
  --line: #2a78d6; --ref: #85847f;
  --good: #0ca30c; --critical: #d03b3b;
}
@media (prefers-color-scheme: dark) {
  :root:where(:not([data-theme="light"])) {
    color-scheme: dark;
    --bg: #121211; --surface: #1a1a19; --border: #2e2e2b;
    --ink: #f0efec; --ink-2: #c3c2b7; --ink-3: #8f8e86; --grid: #2a2a27;
    --over: #e66767; --under: #3987e5; --neutral: #5a5955;
    --line: #3987e5; --ref: #8f8e86;
  }
}
:root[data-theme="dark"] {
  color-scheme: dark;
  --bg: #121211; --surface: #1a1a19; --border: #2e2e2b;
  --ink: #f0efec; --ink-2: #c3c2b7; --ink-3: #8f8e86; --grid: #2a2a27;
  --over: #e66767; --under: #3987e5; --neutral: #5a5955;
  --line: #3987e5; --ref: #8f8e86;
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--ink);
  font: 15px/1.45 "Segoe UI", system-ui, -apple-system, sans-serif; }
main { max-width: 1080px; margin: 0 auto; padding: 28px 16px 48px; }
header { display: flex; flex-wrap: wrap; align-items: end; justify-content: space-between; gap: 12px; margin-bottom: 20px; }
h1 { font-size: 24px; margin: 0; font-weight: 650; letter-spacing: -0.01em; }
.sub { color: var(--ink-2); margin-top: 2px; }
.meta { color: var(--ink-3); font-size: 13px; }
.card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 20px; margin-bottom: 16px; }
h2 { font-size: 16px; margin: 0 0 4px; font-weight: 600; }
.note { color: var(--ink-2); font-size: 13px; margin: 0 0 14px; }
/* Månedsknapperne er diagrammets x-akse: samme margen som diagrammet (L = 58, R = 8), én kolonne pr. søjle */
.months { display: grid; grid-template-columns: repeat(12, 1fr); gap: 4px; padding: 0 8px 0 58px; margin-top: 6px; }
.months button { font: inherit; font-size: 13px; border: 1px solid var(--border); background: var(--surface);
  color: var(--ink-2); border-radius: 999px; padding: 5px 0; cursor: pointer; min-width: 0;
  overflow: hidden; text-overflow: clip; white-space: nowrap; }
.months button[aria-pressed="true"] { background: var(--ink); color: var(--surface); border-color: var(--ink); }
.months button:disabled { opacity: .4; cursor: default; }
.hero { display: flex; flex-wrap: wrap; align-items: center; gap: 10px 18px; }
.hero .big { font-size: 34px; font-weight: 650; letter-spacing: -0.02em; font-variant-numeric: tabular-nums; }
.badge { display: inline-flex; align-items: center; gap: 6px; font-weight: 600; font-size: 14px;
  padding: 4px 10px; border-radius: 999px; border: 1px solid var(--border); }
.badge .dot { width: 10px; height: 10px; border-radius: 50%; }
.tiles { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-top: 18px; }
.tile { border-top: 1px solid var(--border); padding-top: 10px; }
.tile .k { color: var(--ink-2); font-size: 13px; }
.tile .v { font-size: 22px; font-weight: 600; font-variant-numeric: tabular-nums; }
.tile .s { color: var(--ink-3); font-size: 12px; }
.chart { position: relative; }
svg { display: block; overflow: visible; }
svg text { fill: var(--ink-3); font-size: 11px; font-variant-numeric: tabular-nums; }
.legend { display: flex; flex-wrap: wrap; gap: 16px; font-size: 13px; color: var(--ink-2); margin-bottom: 8px; }
.legend i { display: inline-block; width: 14px; height: 10px; border-radius: 2px; margin-right: 6px; vertical-align: -1px; }
.legend i.ln { height: 2px; vertical-align: 3px; }
.legend i.dash { height: 0; border-top: 2px dashed var(--ref); background: none; vertical-align: 3px; }
.tip { position: absolute; pointer-events: none; background: var(--surface); border: 1px solid var(--border);
  border-radius: 8px; padding: 8px 10px; font-size: 13px; box-shadow: 0 4px 16px rgba(0,0,0,.12);
  white-space: nowrap; opacity: 0; transition: opacity .1s; z-index: 2; }
.tip b { display: block; margin-bottom: 2px; }
.tip .row { display: flex; justify-content: space-between; gap: 14px; color: var(--ink-2); }
.tip .row span:last-child { color: var(--ink); font-variant-numeric: tabular-nums; }
table { width: 100%; border-collapse: collapse; font-size: 13px; font-variant-numeric: tabular-nums; }
th, td { padding: 6px 8px; text-align: right; border-bottom: 1px solid var(--grid); }
th:first-child, td:first-child { text-align: left; }
th { color: var(--ink-2); font-weight: 600; position: sticky; top: 0; background: var(--surface); }
.tbl { max-height: 420px; overflow: auto; }
.pos { color: var(--critical); } .neg { color: var(--under); }
.kat { position: relative; display: grid; grid-template-columns: minmax(110px, 190px) 1fr; row-gap: 8px; column-gap: 12px; align-items: center; }
.kat .lbl { font-size: 13px; color: var(--ink-2); text-align: right; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.kat .track { position: relative; height: 26px; cursor: default; }
.kat .zero { position: absolute; top: -4px; bottom: -4px; width: 1px; background: var(--ink-3); }
.kat .bar { position: absolute; top: 4px; height: 18px; min-width: 2px; }
.kat .bar.op { border-radius: 0 4px 4px 0; } .kat .bar.ned { border-radius: 4px 0 0 4px; }
.kat .val { position: absolute; top: 4px; font-size: 12px; line-height: 18px; color: var(--ink); font-variant-numeric: tabular-nums; white-space: nowrap; }
.kat .row:hover .lbl, .kat .track:hover ~ * { color: var(--ink); }
details summary { cursor: pointer; color: var(--ink-2); font-size: 14px; }
@media (max-width: 720px) { .tiles { grid-template-columns: repeat(2, 1fr); } .hero .big { font-size: 28px; } }
</style>
</head>
<body>
<main>
  <header>
    <div>
      <h1>Køkkenets indkøb mod budget</h1>
      <div class="sub" id="kantine"></div>
    </div>
    <div class="meta" id="opdateret"></div>
  </header>

  <section class="card">
    <h2>Pr. måned: har vi købt for meget eller for lidt?</h2>
    <p class="note">Indkøb minus budget. Budget = kuverter × kuvertpris + vareforbrug fra ekstrasalg. Over nul: købt for meget. Under nul: købt for lidt.</p>
    <div class="legend"><span><i style="background:var(--over)"></i>Over budget</span><span><i style="background:var(--under)"></i>Under budget</span><span><i style="background:var(--neutral)"></i>Inden for ±<span id="tol"></span> %</span></div>
    <div class="chart" id="maanedChart"></div>
    <div class="months" id="months" role="group" aria-label="Vælg måned"></div>
  </section>

  <section class="card">
    <div class="hero" id="hero"></div>
    <div class="tiles" id="tiles"></div>
  </section>

  <section class="card">
    <h2 id="katTitel">Hvor kommer forskellen fra?</h2>
    <p class="note">Hvad hver kolonne ville have kostet, hvis måneden havde ramt budgettet med den normale fordeling mellem kolonnerne (<span id="katRef"></span>), sammenlignet med hvad der faktisk er købt. Søjlerne lagt sammen giver månedens samlede forskel.</p>
    <div class="kat" id="katChart"></div>
    <p class="note" id="katKonklusion" style="margin:14px 0 0"></p>
  </section>

  <section class="card">
    <h2 id="dagTitel">Løbende indkøb mod løbende budget</h2>
    <p class="note">Hvor langt indkøbet er nået i måneden, sammenlignet med budgettet for de gæster, der har spist indtil da.</p>
    <div class="legend"><span><i class="ln" style="background:var(--line)"></i>Indkøb (løbende)</span><span><i class="dash"></i>Budget (løbende)</span></div>
    <div class="chart" id="dagChart"></div>
  </section>

  <section class="card">
    <details>
      <summary>Tabel pr. dag</summary>
      <div class="tbl" style="margin-top:12px"><table id="tabel"></table></div>
    </details>
  </section>
</main>

<script>
const D = /*DATA*/null;
const kr = (v, d = 0) => v.toLocaleString("da-DK", { minimumFractionDigits: d, maximumFractionDigits: d }) + " kr.";
const n0 = v => v.toLocaleString("da-DK", { maximumFractionDigits: 0 });
const pct = v => (v > 0 ? "+" : "") + (v * 100).toLocaleString("da-DK", { maximumFractionDigits: 0 }) + " %";
const NS = "http://www.w3.org/2000/svg";
const el = (tag, attrs = {}, parent) => { const e = document.createElementNS(NS, tag);
  for (const k in attrs) e.setAttribute(k, attrs[k]); if (parent) parent.appendChild(e); return e; };

// Månedstal
const M = D.maaneder.map(m => {
  const s = k => m.dage.reduce((a, d) => a + d[k], 0);
  const indkoeb = s("indkoeb"), budget = s("budget"), kuverter = s("kuverter"), vf = s("vareforbrug");
  const diff = indkoeb - budget;
  const kat = (D.kolonner || []).map((_, i) => m.dage.reduce((a, d) => a + d.kat[i], 0));
  return { ...m, indkoeb, budget, kuverter, diff, kat, pr: kuverter ? (indkoeb - vf) / kuverter : null,
           har: indkoeb > 0 || kuverter > 0, rel: budget ? diff / budget : null };
});
const status = m => m.rel === null ? { t: "Intet budget", c: "var(--neutral)", s: "–" }
  : m.rel > D.tolerance ? { t: "Købt for meget", c: "var(--over)", s: "▲" }
  : m.rel < -D.tolerance ? { t: "Købt for lidt", c: "var(--under)", s: "▼" }
  : { t: "På budget", c: "var(--neutral)", s: "●" };

document.getElementById("kantine").textContent = (D.kantine ? D.kantine + " · " : "") + D.aar + " · kuvertpris " + kr(D.kuvertpris, 2);
document.getElementById("opdateret").textContent = "Opdateret " + D.opdateret;
document.getElementById("tol").textContent = Math.round(D.tolerance * 100);

function tip(box) { let t = box.querySelector(".tip"); if (!t) { t = document.createElement("div"); t.className = "tip"; box.appendChild(t); } return t; }
function showTip(box, html, x, y) { const t = tip(box); t.innerHTML = html; t.style.opacity = 1;
  const w = t.offsetWidth, bw = box.clientWidth; t.style.left = Math.min(Math.max(x - w / 2, 0), bw - w) + "px"; t.style.top = (y - t.offsetHeight - 10) + "px"; }
function hideTip(box) { tip(box).style.opacity = 0; }
function ticks(min, max, n = 5) { const span = max - min || 1, step0 = span / n, mag = 10 ** Math.floor(Math.log10(step0));
  const step = [1, 2, 2.5, 5, 10].map(f => f * mag).find(s => span / s <= n) || mag * 10; const out = [];
  for (let v = Math.floor(min / step) * step; v < max + step - 1e-9; v += step) out.push(+v.toFixed(6)); return out; }

// Diagram 1: afvigelse pr. måned (divergerende søjler)
function maanedChart() {
  const box = document.getElementById("maanedChart"); box.innerHTML = "";
  const W = Math.max(box.clientWidth, 300), H = 240, L = 58, R = 8, T = 12, B = 6;
  const vals = M.map(m => m.har ? m.diff : 0);
  let lo = Math.min(0, ...vals), hi = Math.max(0, ...vals); const tk = ticks(lo, hi); lo = Math.min(lo, tk[0]); hi = Math.max(hi, tk[tk.length - 1]);
  const y = v => T + (hi - v) / (hi - lo || 1) * (H - T - B);
  const svg = el("svg", { viewBox: `0 0 ${W} ${H}`, width: W, height: H, role: "img", "aria-label": "Afvigelse fra budget pr. måned" }, box);
  tk.forEach(v => { el("line", { x1: L, x2: W - R, y1: y(v), y2: y(v), stroke: v === 0 ? "var(--ink-3)" : "var(--grid)", "stroke-width": 1 }, svg);
    el("text", { x: L - 8, y: y(v) + 4, "text-anchor": "end" }, svg).textContent = n0(v); });
  const bw = (W - L - R) / 12;
  M.forEach((m, i) => {
    const x = L + i * bw, cx = x + bw / 2;
    const knap = document.querySelector(`#months button[data-nr="${m.nr}"]`);   // knapperne er x-aksens etiketter
    if (knap) knap.textContent = bw >= 84 ? m.navn : bw >= 34 ? m.navn.slice(0, 3) : m.navn.slice(0, 1);
    if (!m.har) return;
    const st = status(m), y0 = y(0), y1 = y(m.diff), h = Math.max(Math.abs(y1 - y0), 1), up = m.diff >= 0;
    const w = Math.min(bw * 0.56, 44), r = Math.min(4, h / 2);
    // Afrundede ender væk fra nullinjen
    const top = up ? y1 : y0, bot = up ? y0 : y1, xl = cx - w / 2, xr = cx + w / 2;
    const d = up ? `M${xl},${bot} V${top + r} Q${xl},${top} ${xl + r},${top} H${xr - r} Q${xr},${top} ${xr},${top + r} V${bot} Z`
                 : `M${xl},${top} V${bot - r} Q${xl},${bot} ${xl + r},${bot} H${xr - r} Q${xr},${bot} ${xr},${bot - r} V${top} Z`;
    el("path", { d, fill: st.c }, svg);
    if (m.nr === valgt) el("rect", { x: x + 2, y: T, width: bw - 4, height: H - T - B, fill: "none", stroke: "var(--ink-3)", "stroke-dasharray": "3 3", rx: 6 }, svg);
    const hit = el("rect", { x, y: T, width: bw, height: H - T - B, fill: "transparent", style: "cursor:pointer" }, svg);
    hit.addEventListener("mousemove", e => { const b = box.getBoundingClientRect();
      showTip(box, `<b>${m.navn}</b><div class="row"><span>Indkøb</span><span>${kr(m.indkoeb)}</span></div><div class="row"><span>Budget</span><span>${kr(m.budget)}</span></div><div class="row"><span>Forskel</span><span>${(m.diff > 0 ? "+" : "") + kr(m.diff)}${m.rel !== null ? " (" + pct(m.rel) + ")" : ""}</span></div><div class="row"><span>Kuverter</span><span>${n0(m.kuverter)}</span></div><div class="row"><span>Status</span><span>${st.s} ${st.t}</span></div>`,
        e.clientX - b.left, e.clientY - b.top); });
    hit.addEventListener("mouseleave", () => hideTip(box));
    hit.addEventListener("click", () => vaelg(m.nr));
  });
}

// Diagram 2: løbende indkøb og budget i den valgte måned
function dagChart(m) {
  const box = document.getElementById("dagChart"); box.innerHTML = "";
  let a = 0, b = 0; const pts = m.dage.map(d => ({ dag: d.dag, d, ind: a += d.indkoeb, bud: b += d.budget }));
  const W = Math.max(box.clientWidth, 300), H = 280, L = 58, R = 12, T = 12, B = 26;
  const hi0 = Math.max(1, ...pts.map(p => Math.max(p.ind, p.bud))); const tk = ticks(0, hi0); const hi = Math.max(hi0, tk[tk.length - 1]);
  const n = pts.length, x = i => L + (n > 1 ? i / (n - 1) : 0.5) * (W - L - R), y = v => T + (1 - v / hi) * (H - T - B);
  const svg = el("svg", { viewBox: `0 0 ${W} ${H}`, width: W, height: H, role: "img", "aria-label": "Løbende indkøb og budget i " + m.navn }, box);
  tk.forEach(v => { el("line", { x1: L, x2: W - R, y1: y(v), y2: y(v), stroke: v === 0 ? "var(--ink-3)" : "var(--grid)" }, svg);
    el("text", { x: L - 8, y: y(v) + 4, "text-anchor": "end" }, svg).textContent = n0(v); });
  pts.forEach((p, i) => { if (p.dag === 1 || p.dag % 5 === 0) el("text", { x: x(i), y: H - 8, "text-anchor": "middle" }, svg).textContent = p.dag; });
  const path = k => pts.map((p, i) => (i ? "L" : "M") + x(i).toFixed(1) + "," + y(p[k]).toFixed(1)).join(" ");
  el("path", { d: path("bud"), fill: "none", stroke: "var(--ref)", "stroke-width": 2, "stroke-dasharray": "6 5" }, svg);
  el("path", { d: path("ind"), fill: "none", stroke: "var(--line)", "stroke-width": 2, "stroke-linejoin": "round" }, svg);
  const last = pts[pts.length - 1];
  if (last) { const st = status(M[m.nr - 1]);
    el("text", { x: W - R, y: y(last.ind) + (last.ind >= last.bud ? -8 : 16), "text-anchor": "end", style: "fill:var(--ink-2);font-size:12px" }, svg).textContent = "Indkøb " + kr(last.ind);
    el("text", { x: W - R, y: y(last.bud) + (last.ind >= last.bud ? 16 : -8), "text-anchor": "end", style: "fill:var(--ink-3);font-size:12px" }, svg).textContent = "Budget " + kr(last.bud); }
  const cross = el("line", { y1: T, y2: H - B, stroke: "var(--ink-3)", "stroke-width": 1, opacity: 0 }, svg);
  const dot1 = el("circle", { r: 5, fill: "var(--line)", stroke: "var(--surface)", "stroke-width": 2, opacity: 0 }, svg);
  const dot2 = el("circle", { r: 5, fill: "var(--ref)", stroke: "var(--surface)", "stroke-width": 2, opacity: 0 }, svg);
  const hit = el("rect", { x: L, y: T, width: W - L - R, height: H - T - B, fill: "transparent" }, svg);
  hit.addEventListener("mousemove", e => {
    const bb = svg.getBoundingClientRect(), sx = (e.clientX - bb.left) / bb.width * W;
    const i = Math.max(0, Math.min(n - 1, Math.round((sx - L) / ((W - L - R) / Math.max(n - 1, 1))))), p = pts[i];
    [cross, dot1, dot2].forEach(o => o.setAttribute("opacity", 1));
    cross.setAttribute("x1", x(i)); cross.setAttribute("x2", x(i));
    dot1.setAttribute("cx", x(i)); dot1.setAttribute("cy", y(p.ind)); dot2.setAttribute("cx", x(i)); dot2.setAttribute("cy", y(p.bud));
    const diff = p.ind - p.bud, b = box.getBoundingClientRect();
    showTip(box, `<b>${p.dag}. ${m.navn.toLowerCase()}</b><div class="row"><span>Kuverter i dag</span><span>${n0(p.d.kuverter)}</span></div><div class="row"><span>Indkøb i dag</span><span>${kr(p.d.indkoeb)}</span></div><div class="row"><span>Indkøb løbende</span><span>${kr(p.ind)}</span></div><div class="row"><span>Budget løbende</span><span>${kr(p.bud)}</span></div><div class="row"><span>Forskel</span><span>${(diff > 0 ? "+" : "") + kr(diff)}</span></div>`,
      x(i) / W * b.width, y(Math.max(p.ind, p.bud)) / H * b.height);
  });
  hit.addEventListener("mouseleave", () => { [cross, dot1, dot2].forEach(o => o.setAttribute("opacity", 0)); hideTip(box); });
}

function tabel(m) {
  const t = document.getElementById("tabel");
  let html = "<thead><tr><th>Dag</th><th>Kuverter</th><th>Indkøb</th><th>Budget</th><th>Forskel</th><th>Kr. pr. kuvert</th></tr></thead><tbody>";
  m.dage.forEach(d => { const diff = d.indkoeb - d.budget, pr = d.kuverter ? (d.indkoeb - d.vareforbrug) / d.kuverter : null;
    if (!d.indkoeb && !d.kuverter && !d.budget) return;
    html += `<tr><td>${d.dag}.</td><td>${n0(d.kuverter)}</td><td>${kr(d.indkoeb, 2)}</td><td>${kr(d.budget, 2)}</td><td class="${diff > 0 ? "pos" : diff < 0 ? "neg" : ""}">${(diff > 0 ? "+" : "") + kr(diff, 2)}</td><td>${pr === null ? "–" : kr(pr, 2)}</td></tr>`; });
  const mm = M[m.nr - 1];
  html += `</tbody><tfoot><tr><th>I alt</th><th>${n0(mm.kuverter)}</th><th>${kr(mm.indkoeb, 2)}</th><th>${kr(mm.budget, 2)}</th><th>${(mm.diff > 0 ? "+" : "") + kr(mm.diff, 2)}</th><th>${mm.pr === null ? "–" : kr(mm.pr, 2)}</th></tr></tfoot>`;
  t.innerHTML = html;
}

// Normal fordeling mellem kolonnerne: hele måneder (alle med data undtagen indeværende)
const nuMd = new Date().getMonth() + 1;
const REF = M.filter(m => m.har && m.indkoeb > 0 && m.nr !== nuMd);
const refInd = REF.reduce((a, m) => a + m.indkoeb, 0);
const ANDEL = D.kolonner.map((_, i) => refInd ? REF.reduce((a, m) => a + m.kat[i], 0) / refInd : 0);
const kortNavn = s => s.replace(/\s*\(.*\)\s*$/, "");

function katChart(m) {
  const box = document.getElementById("katChart"); box.innerHTML = "";
  document.getElementById("katTitel").textContent = "Hvor kommer forskellen fra? – " + m.navn;
  document.getElementById("katRef").textContent = REF.length
    ? "beregnet ud fra " + (REF.length === 1 ? REF[0].navn : REF[0].navn.toLowerCase() + "–" + REF[REF.length - 1].navn.toLowerCase())
    : "ingen hele måneder at sammenligne med endnu";
  const konk = document.getElementById("katKonklusion");
  if (!m.har || !m.budget) { konk.textContent = "Ingen data for " + m.navn.toLowerCase() + " endnu."; return; }
  const rk = D.kolonner.map((navn, i) => { const brugt = m.kat[i], ved = m.budget * ANDEL[i];
    return { navn, brugt, ved, diff: brugt - ved, andel: m.indkoeb ? brugt / m.indkoeb : 0, normal: ANDEL[i] }; })
    .filter(r => r.brugt || r.ved).sort((a, b) => b.diff - a.diff);
  // Nullinjen placeres efter forholdet mellem plus og minus; samme kr. giver samme længde på begge sider
  const maxPos = Math.max(0, ...rk.map(r => r.diff)), maxNeg = Math.max(0, ...rk.map(r => -r.diff));
  const z = 0.08 + 0.84 * (maxNeg / ((maxPos + maxNeg) || 1));
  const span = 0.7;    // resten er plads til beløbet ved enden af søjlen
  const skala = span * Math.min(maxPos ? (1 - z) / maxPos : Infinity, maxNeg ? z / maxNeg : Infinity);
  rk.forEach(r => {
    const lbl = document.createElement("div"); lbl.className = "lbl"; lbl.textContent = kortNavn(r.navn); lbl.title = r.navn;
    const tr = document.createElement("div"); tr.className = "track";
    const w = Math.abs(r.diff) * (isFinite(skala) ? skala : 0) * 100;
    const bar = document.createElement("div"); bar.className = "bar " + (r.diff >= 0 ? "op" : "ned");
    bar.style.background = Math.abs(r.diff) < m.budget * 0.005 ? "var(--neutral)" : r.diff > 0 ? "var(--over)" : "var(--under)";
    bar.style.width = w + "%"; bar.style.left = (r.diff >= 0 ? z * 100 : z * 100 - w) + "%";
    const val = document.createElement("div"); val.className = "val"; val.textContent = (r.diff > 0 ? "+" : "") + kr(r.diff);
    if (r.diff >= 0) val.style.left = `calc(${z * 100 + w}% + 6px)`; else val.style.right = `calc(${100 - z * 100 + w}% + 6px)`;
    const zero = document.createElement("div"); zero.className = "zero"; zero.style.left = z * 100 + "%";
    tr.append(zero, bar, val);
    tr.addEventListener("mousemove", e => { const b = box.getBoundingClientRect();
      showTip(box, `<b>${r.navn}</b><div class="row"><span>Købt</span><span>${kr(r.brugt)}</span></div><div class="row"><span>Ved budget</span><span>${kr(r.ved)}</span></div><div class="row"><span>Forskel</span><span>${(r.diff > 0 ? "+" : "") + kr(r.diff)}</span></div><div class="row"><span>Andel af indkøb</span><span>${pct(r.andel).replace("+", "")} (normalt ${pct(r.normal).replace("+", "")})</span></div>`,
        e.clientX - b.left, e.clientY - b.top); });
    tr.addEventListener("mouseleave", () => hideTip(box));
    box.append(lbl, tr);
  });
  const top = rk.filter(r => r.diff > 0).slice(0, 2);
  konk.textContent = m.diff > 0 && top.length
    ? `${m.navn} ligger ${kr(m.diff)} over budget. Største bidrag: ` + top.map(r => `${kortNavn(r.navn)} (${(r.diff > 0 ? "+" : "") + kr(r.diff)})`).join(" og ") + "."
    : m.diff < 0 ? `${m.navn} ligger ${kr(-m.diff)} under budget.` + (top.length ? ` Over det normale: ` + top.map(r => `${kortNavn(r.navn)} (+${kr(r.diff)})`).join(" og ") + "." : "")
    : `${m.navn} ligger på budget.`;
}

function hero(m) {
  const st = status(m);
  document.getElementById("hero").innerHTML = m.har
    ? `<div><div class="sub">${m.navn} ${D.aar}</div><div class="big">${(m.diff > 0 ? "+" : "") + kr(m.diff)}</div></div>
       <span class="badge"><span class="dot" style="background:${st.c}"></span>${st.s} ${st.t}${m.rel !== null ? " · " + pct(m.rel) : ""}</span>`
    : `<div><div class="sub">${m.navn} ${D.aar}</div><div class="big">Ingen data endnu</div></div>`;
  const prDiff = m.pr === null ? "" : `${(m.pr - D.kuvertpris > 0 ? "+" : "") + kr(m.pr - D.kuvertpris, 2)} mod kuvertprisen`;
  document.getElementById("tiles").innerHTML = [
    ["Indkøb", kr(m.indkoeb), "kolonnerne BC catering – Carlsberg"],
    ["Budget", kr(m.budget), "kuverter × " + kr(D.kuvertpris, 2) + " + vareforbrug"],
    ["Kuverter", n0(m.kuverter), "fra Kanpla / manuelt"],
    ["Kr. pr. kuvert", m.pr === null ? "–" : kr(m.pr, 2), prDiff]
  ].map(([k, v, s]) => `<div class="tile"><div class="k">${k}</div><div class="v">${v}</div><div class="s">${s}</div></div>`).join("");
}

let valgt = 0;
function vaelg(nr) {
  valgt = nr; const m = M[nr - 1];
  document.querySelectorAll("#months button").forEach(b => b.setAttribute("aria-pressed", +b.dataset.nr === nr));
  document.getElementById("dagTitel").textContent = "Løbende indkøb mod løbende budget – " + m.navn;
  hero(m); katChart(m); dagChart(m); tabel(m); maanedChart();
}
const mb = document.getElementById("months");
M.forEach(m => { const b = document.createElement("button"); b.textContent = m.navn; b.title = m.navn; b.dataset.nr = m.nr; b.disabled = !m.har;
  b.addEventListener("click", () => vaelg(m.nr)); mb.appendChild(b); });
const nu = new Date().getMonth() + 1;
let rt, bredde = 0; new ResizeObserver(e => { const w = e[0].contentRect.width; if (Math.abs(w - bredde) < 2) return; bredde = w;
  clearTimeout(rt); rt = setTimeout(() => valgt && vaelg(valgt), 100); }).observe(document.getElementById("dagChart"));
vaelg((M.find(m => m.nr === nu && m.har) || [...M].reverse().find(m => m.har) || M[0]).nr);
</script>
</body>
</html>
"""


def main():
    from ordre_til_ark import ARK, aabn_ark
    xl, wb, egen = aabn_ark(ARK)
    try:
        sti = skriv_html(hent_data(wb))
        print(f"Rapport skrevet: {sti}")
    finally:
        if egen:
            wb.Close(False)
            xl.Quit()


if __name__ == "__main__":
    main()
