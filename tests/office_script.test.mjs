// Tester Office Scriptet i Node med en lille attrap af Excel-tabellerne.
// Kør:  node --test tests/office_script.test.mjs
import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import test from "node:test";
import assert from "node:assert/strict";

const ROD = new URL("../", import.meta.url);
const kilde = readFileSync(new URL("office-scripts/IndlaesOrdrekvittering.ts", ROD), "utf8");
const js = stripTypeScriptTypes(kilde.replace(/^interface \w+ \{[\s\S]*?^\}/gm, ""));
const { main, parseOrdrekvittering } = new Function(js + "\nreturn { main, parseOrdrekvittering };")();
const MAIL = readFileSync(new URL("eksempler/indbakke/ordrekvittering_31525653.html", ROD), "utf8");

function tabel(overskrifter, raekker = [overskrifter.map(() => "")]) {
  const data = raekker.map((r) => [...r]);
  const omraade = (rows) => ({
    getValues: () => rows().map((r) => [...r]),
    getRow: (i) => ({ setValues: ([v]) => { data[i] = [...v]; } }),
    setNumberFormat: () => {},
  });
  return {
    data,
    getHeaderRowRange: () => ({ getValues: () => [overskrifter] }),
    getRangeBetweenHeaderAndTotal: () => omraade(() => data),
    getColumnByName: (n) => ({ getRangeBetweenHeaderAndTotal: () => omraade(() => data.map((r) => [r[overskrifter.indexOf(n)]])) }),
    addRows: (_i, v) => data.push(...v.map((r) => [...r])),
  };
}

function arbejdsbog() {
  const t = {
    Ordrer: tabel(["Ordre", "Leveringsdato", "Leverandør", "Kunde", "Bruger", "Leveringsbemærkning", "Kolonne",
      "Antal linjer", "Total", "Kilde", "Indlæst", "Advarsler"]),
    Ordrelinjer: tabel(["Ordre", "Leveringsdato", "Leverandør", "Kunde", "Bruger", "Leveringsbemærkning",
      "Varenr", "Vare", "Antal", "Enhed", "Enhedspris", "Linjepris"]),
    Kolonneregler: tabel(["Leverandør", "Genkend i mail", "Kunde indeholder", "Kolonne"], [
      ["AB Catering København", "AB Catering", "FRUGT", "Frugt"],
      ["AB Catering København", "AB Catering", "", "BC catering"],
    ]),
  };
  return { t, getTable: (n) => t[n] };
}

test("læser den rigtige ordrekvittering", () => {
  const o = parseOrdrekvittering(MAIL);
  assert.equal(o.ordre, "31525653");
  assert.equal(o.leveringsdato, "27/10/2026");
  assert.equal(o.kunde, "39-69505002, FORENEDE SERVICE FRUGT");
  assert.equal(o.bruger, "kaso@kokkeneskoekken.dk");
  assert.equal(o.linjer.length, 5);
  assert.deepEqual(o.linjer[0], { varenr: "224918", vare: "Æble Ingrid Marie DK Øko", antal: 20, enhed: "STK", enhedspris: 3.42, linjepris: 68.4 });
  assert.equal(o.total, 386.23);
});

test("skriver ordren ind i tabellerne, og springer dubletter over", () => {
  const wb = arbejdsbog();
  const r = main(wb, MAIL, "Ordrekvittering", "2026-10-27T08:00:00Z");
  assert.equal(r.status, "indlæst");
  assert.equal(r.kolonne, "Frugt");
  assert.equal(r.advarsler, "");
  assert.equal(wb.t.Ordrer.data.length, 1);
  const [ordre, dato, lev, , , , kol, antal, total] = wb.t.Ordrer.data[0];
  assert.deepEqual([ordre, dato, lev, kol, antal, total], ["31525653", 46322, "AB Catering København", "Frugt", 5, 386.23]);
  assert.equal(wb.t.Ordrelinjer.data.length, 5);
  assert.equal(wb.t.Ordrelinjer.data[4][7], "Tatsoi Mini Øko");

  assert.equal(main(wb, MAIL, "Ordrekvittering", "").status, "dublet");
  assert.equal(wb.t.Ordrer.data.length, 1);
});

test("HTML-entiteter og indlejrede layout-tabeller", () => {
  const html = `<table><tr><td><table>
    <tr><td>Ordre:</td><td>999001</td></tr>
    <tr><td>Kunde:</td><td>39-1, FORENEDE SERVICE KANTINE</td></tr>
    <tr><td>Leveringsdato:</td><td>03/11/2026</td></tr></table>
    <table><tr><td>510233</td><td>Rugbr&oslash;d &amp; Gr&#248;nt</td><td>2</td><td>STK</td><td>1.028,50</td><td>2.057,00</td></tr>
    <tr><td></td><td>Total</td><td>2.057,00</td></tr></table>
    </td></tr></table><p>AB Catering K&oslash;benhavn</p>`;
  const wb = arbejdsbog();
  const r = main(wb, html, "Ordrekvittering", "");
  assert.equal(r.status, "indlæst");
  assert.equal(r.kolonne, "BC catering");
  assert.equal(wb.t.Ordrelinjer.data[0][7], "Rugbrød & Grønt");
  assert.equal(r.total, 2057);
});

test("ukendt leverandør og mail uden ordre", () => {
  const wb = arbejdsbog();
  const ukendt = MAIL.replace("AB Catering København", "Anden Grossist");
  assert.equal(main(wb, ukendt, "", "").status, "ukendt leverandør");
  assert.equal(main(arbejdsbog(), "<p>Hej, tak for sidst</p>", "", "").status, "fejl");
});
