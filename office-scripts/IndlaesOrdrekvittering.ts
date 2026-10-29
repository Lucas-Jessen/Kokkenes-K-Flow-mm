/**
 * Office Script: Indlæs ordrekvittering
 *
 * Kaldes af Power Automate for hver ny mail i indkoebkk@forenede.dk.
 * Læser mailens HTML og skriver ordren ind i tabellerne "Ordrer" og "Ordrelinjer" i K-flow.xlsx.
 * Hvilken kolonne (konto) ordren tæller i, styres af tabellen "Kolonneregler" i samme fil.
 * En ordre, der allerede findes (samme ordrenummer), skrives ikke igen.
 *
 * Returnerer status til flowet: "indlæst" | "dublet" | "ukendt leverandør" | "fejl".
 */

interface Varelinje {
  varenr: string;
  vare: string;
  antal: number;
  enhed: string;
  enhedspris: number;
  linjepris: number;
}

interface Ordre {
  ordre: string;
  leveringsdato: string; // dd/mm/åååå som i mailen
  kunde: string;
  bruger: string;
  leveringsbemaerkning: string;
  linjer: Varelinje[];
  total: number;
  tekst: string;
}

interface Regel {
  leverandoer: string;
  genkend: string;
  kundeIndeholder: string;
  kolonne: string;
}

interface Resultat {
  status: string;
  ordre: string;
  leverandoer: string;
  kolonne: string;
  total: number;
  linjer: number;
  advarsler: string;
}

function main(workbook: ExcelScript.Workbook, mailHtml: string, emne: string, modtaget: string): Resultat {
  const ordre = parseOrdrekvittering(mailHtml);
  const advarsler = valider(ordre);
  const regler = laesRegler(workbook);
  const match = findRegel(regler, ordre);

  const res: Resultat = {
    status: "", ordre: ordre.ordre, leverandoer: match ? match.leverandoer : "Ukendt",
    kolonne: match ? match.kolonne : "", total: ordre.total, linjer: ordre.linjer.length, advarsler: advarsler.join("; "),
  };
  if (!ordre.ordre || ordre.linjer.length === 0) {
    res.status = "fejl";
    return res;
  }

  const ordreTabel = workbook.getTable("Ordrer");
  const linjeTabel = workbook.getTable("Ordrelinjer");
  if (!ordreTabel || !linjeTabel) {
    res.status = "fejl";
    res.advarsler = "Tabellerne Ordrer/Ordrelinjer findes ikke i arbejdsbogen";
    return res;
  }
  const eksisterende = ordreTabel.getColumnByName("Ordre").getRangeBetweenHeaderAndTotal().getValues();
  if (eksisterende.some((r) => String(r[0]) === ordre.ordre)) {
    res.status = "dublet";
    return res;
  }
  if (!match) advarsler.push("Ukendt leverandør – tilføj den i Kolonneregler");

  const dato = datoTilSerial(ordre.leveringsdato);
  tilfoej(ordreTabel, [{
    "Ordre": ordre.ordre, "Leveringsdato": dato, "Leverandør": res.leverandoer, "Kunde": ordre.kunde,
    "Bruger": ordre.bruger, "Leveringsbemærkning": ordre.leveringsbemaerkning, "Kolonne": res.kolonne,
    "Antal linjer": ordre.linjer.length, "Total": ordre.total, "Kilde": `Mail: ${emne}`,
    "Indlæst": modtaget, "Advarsler": advarsler.join("; "),
  }]);
  tilfoej(linjeTabel, ordre.linjer.map((l) => ({
    "Ordre": ordre.ordre, "Leveringsdato": dato, "Leverandør": res.leverandoer, "Kunde": ordre.kunde,
    "Bruger": ordre.bruger, "Leveringsbemærkning": ordre.leveringsbemaerkning, "Varenr": l.varenr, "Vare": l.vare,
    "Antal": l.antal, "Enhed": l.enhed, "Enhedspris": l.enhedspris, "Linjepris": l.linjepris,
  })));
  for (const t of [ordreTabel, linjeTabel]) {
    t.getColumnByName("Leveringsdato").getRangeBetweenHeaderAndTotal().setNumberFormat("dd-mm-yyyy");
  }

  res.advarsler = advarsler.join("; ");
  res.status = match ? "indlæst" : "ukendt leverandør";
  return res;
}

// ---------- Skrivning til Excel ----------

function tilfoej(tabel: ExcelScript.Table, raekker: { [kolonne: string]: string | number }[]) {
  const overskrifter = tabel.getHeaderRowRange().getValues()[0].map((v) => String(v));
  const vaerdier = raekker.map((r) => overskrifter.map((h) => (h in r ? r[h] : "")));
  // En ny skabelon har én tom række i tabellen – den fyldes først
  const krop = tabel.getRangeBetweenHeaderAndTotal();
  const foerste = krop.getValues();
  if (foerste.length === 1 && foerste[0].every((v) => v === "")) {
    krop.getRow(0).setValues([vaerdier[0]]);
    vaerdier.shift();
  }
  if (vaerdier.length > 0) tabel.addRows(-1, vaerdier);
}

function laesRegler(workbook: ExcelScript.Workbook): Regel[] {
  const t = workbook.getTable("Kolonneregler");
  if (!t) return [];
  const h = t.getHeaderRowRange().getValues()[0].map((v) => String(v));
  return t.getRangeBetweenHeaderAndTotal().getValues()
    .filter((r) => String(r[h.indexOf("Leverandør")]).trim() !== "")
    .map((r) => ({
      leverandoer: String(r[h.indexOf("Leverandør")]).trim(),
      genkend: String(r[h.indexOf("Genkend i mail")]).trim(),
      kundeIndeholder: String(r[h.indexOf("Kunde indeholder")]).trim(),
      kolonne: String(r[h.indexOf("Kolonne")]).trim(),
    }));
}

/** Første regel hvor leverandøren genkendes i mailen og (hvis udfyldt) kunden indeholder teksten. */
function findRegel(regler: Regel[], o: Ordre): Regel | undefined {
  const tekst = o.tekst.toLowerCase();
  return regler.find((r) =>
    tekst.includes((r.genkend || r.leverandoer).toLowerCase()) &&
    (!r.kundeIndeholder || o.kunde.toLowerCase().includes(r.kundeIndeholder.toLowerCase())));
}

function datoTilSerial(ddmmaaaa: string): number | string {
  const m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(ddmmaaaa.trim());
  if (!m) return ddmmaaaa;
  return (Date.UTC(+m[3], +m[2] - 1, +m[1]) - Date.UTC(1899, 11, 30)) / 86400000;
}

function valider(o: Ordre): string[] {
  const adv: string[] = [];
  if (!o.ordre) adv.push("Ordrenummer ikke fundet");
  if (!o.leveringsdato) adv.push("Leveringsdato ikke fundet");
  if (o.linjer.length === 0) adv.push("Ingen varelinjer fundet");
  const sum = Math.round(o.linjer.reduce((a, l) => a + l.linjepris, 0) * 100) / 100;
  if (Math.abs(sum - o.total) > 0.05) adv.push(`Sum af linjer (${sum}) matcher ikke total (${o.total})`);
  return adv;
}

// ---------- Læsning af mailen (samme logik som kkflow/parse_ordre.py) ----------

const HOVEDFELTER: { [k: string]: string } = {
  "ordre": "ordre", "leveringsdato": "leveringsdato", "kunde": "kunde",
  "bruger": "bruger", "leveringsbemærkning": "leveringsbemaerkning",
};

function tal(s: string): number {
  const t = s.replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
  return t ? parseFloat(t) : 0;
}

function afkod(s: string): string {
  const navne: { [k: string]: string } = {
    nbsp: " ", amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'",
    aelig: "æ", oslash: "ø", aring: "å", AElig: "Æ", Oslash: "Ø", Aring: "Å",
  };
  return s
    .replace(/&#x([0-9a-f]+);/gi, (_m: string, h: string) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_m: string, d: string) => String.fromCharCode(parseInt(d, 10)))
    .replace(/&([a-z]+);/gi, (m: string, n: string) => (n in navne ? navne[n] : m));
}

/** Gennemløber HTML'en og samler tabelrækker (lister af celletekster) + ren tekst. */
function htmlTilRaekker(html: string): { raekker: string[][]; tekst: string } {
  const ren = html
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<(style|script|head)\b[\s\S]*?<\/\1>/gi, "");
  const raekker: string[][] = [];
  const tekst: string[] = [];
  let raekke: string[] | null = null;
  let celle: string[] | null = null;
  const re = /<(\/?)([a-z0-9]+)[^>]*>|([^<]+)/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(ren)) !== null) {
    if (m[3] !== undefined) {
      const d = afkod(m[3]);
      if (celle) celle.push(d);
      tekst.push(d);
      continue;
    }
    const slut = m[1] === "/";
    const tag = m[2].toLowerCase();
    if (!slut && tag === "tr") raekke = [];
    else if (!slut && (tag === "td" || tag === "th")) celle = [];
    else if (tag === "br" || tag === "p" || tag === "div") tekst.push("\n");
    else if (slut && (tag === "td" || tag === "th") && celle) {
      if (raekke) raekke.push(celle.join("").replace(/\s+/g, " ").trim());
      celle = null;
      tekst.push("\n");
    } else if (slut && tag === "tr" && raekke) {
      if (raekke.some((c) => c !== "")) raekker.push(raekke);
      raekke = null;
    }
  }
  return { raekker, tekst: tekst.join("") };
}

function parseOrdrekvittering(html: string): Ordre {
  const o: Ordre = { ordre: "", leveringsdato: "", kunde: "", bruger: "", leveringsbemaerkning: "", linjer: [], total: NaN, tekst: "" };
  const erHtml = /<\/?(td|tr|table|br|p|div)\b/i.test(html);
  const { raekker, tekst } = erHtml ?htmlTilRaekker(html) : { raekker: [] as string[][], tekst: html };
  o.tekst = tekst;

  for (const r of raekker) {
    const celler = r.filter((c) => c !== "");
    if (celler.length === 0) continue;
    const noegle = celler[0].replace(/:$/, "").trim().toLowerCase();
    if (noegle in HOVEDFELTER && r[0].trim().endsWith(":")) {
      (o as unknown as { [k: string]: string })[HOVEDFELTER[noegle]] = celler.length > 1 ? celler[1] : "";
    } else if (/^\d{4,}$/.test(celler[0]) && celler.length >= 6) {
      o.linjer.push({
        varenr: celler[0], vare: celler[1], antal: tal(celler[2]), enhed: celler[3],
        enhedspris: tal(celler[4]), linjepris: tal(celler[5]),
      });
    } else if (celler[0].toLowerCase() === "total" && celler.length >= 2) {
      o.total = tal(celler[celler.length - 1]);
    }
  }
  if (o.linjer.length === 0) fraTekst(tekst, o);
  if (isNaN(o.total)) o.total = Math.round(o.linjer.reduce((a, l) => a + l.linjepris, 0) * 100) / 100;
  return o;
}

/** Fallback når mailen ikke har tabeller: én varelinje pr. tekstlinje. */
function fraTekst(tekst: string, o: Ordre) {
  const linjeRe = /^(\d{4,})\s+(.+?)\s+([\d.,]+)\s+([A-Za-zÆØÅæøå]+)\s+([\d.]*\d,\d{2})\s+([\d.]*\d,\d{2})\s*$/;
  for (const raa of tekst.split(/\r?\n/)) {
    const linje = raa.trim();
    let m = /^(Ordre|Leveringsdato|Kunde|Bruger|Leveringsbemærkning):\s*(.*)$/i.exec(linje);
    if (m) {
      (o as unknown as { [k: string]: string })[HOVEDFELTER[m[1].toLowerCase()]] = m[2].trim();
      continue;
    }
    m = linjeRe.exec(linje);
    if (m) {
      o.linjer.push({ varenr: m[1], vare: m[2], antal: tal(m[3]), enhed: m[4], enhedspris: tal(m[5]), linjepris: tal(m[6]) });
      continue;
    }
    m = /^Total\s+([\d.]*\d,\d{2})$/i.exec(linje);
    if (m) o.total = tal(m[1]);
  }
}
