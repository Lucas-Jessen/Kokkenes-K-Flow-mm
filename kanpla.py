"""Henter spisende gæster fra Kanpla (Data Export API) og skriver dem i fanen "Kanpla" i Kokkenes k ark.xlsx.

Månedsfanernes "Kuverter" henter selv antallet fra Kanpla-fanen med formler (pr. dato).
Står der et tal i "Kuvert manuel", bruges det i stedet for Kanpla den dag.
Opdateringen kører automatisk sammen med ordre_til_ark.py (hver time).

Nøglen (export & billing) ligger i kanpla_api_noegle.txt ved siden af dette script.
Filen er i .gitignore – repoet er offentligt, så nøglen må aldrig committes.

  python kanpla.py --opsaet                              # opret Kanpla-fanen + Kuverter-formler (kun én gang)
  python kanpla.py --opdater                             # hent gæster ind i arket nu
  python kanpla.py 2026-10-26 2026-10-30                 # vis gæster pr. dag pr. lokation (alle lokationer)
  python kanpla.py 2026-10-26 2026-10-30 --lokation "TDC Sletvej"
  python kanpla.py 2026-10-26 2026-10-30 --moduler       # vis alle moduler (også POS/café)
"""
import argparse
import json
import urllib.request
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

ROD = Path(__file__).resolve().parent
NOEGLE_FIL = ROD / "kanpla_api_noegle.txt"
EXPORT_URL = "https://export.api.kanpla.io/v3/export"
GAESTE_MODUL = "registr"          # moduler hvis navn indeholder dette tæller som spisende gæster (frokostregistrering)

# Køkkenet arket hører til (Borupvang hedder "HQ Ballerup" i Kanpla)
LOKATION = "HQ Ballerup"
LOKATION_ID = "y9kKlZ6E2COdaB2e3Ux9"
DAGE_TILBAGE = 14                 # registreringer kan rettes bagud, så de seneste dage hentes igen hver gang

FANE = "Kanpla"
KOLONNER = ["Dato", "Måned", "Dag", "Lokation", "Gæster", "Hentet"]
MDR = ["Januar", "Februar", "Marts", "April", "Maj", "Juni", "Juli", "August",
       "September", "Oktober", "November", "December"]


# ---------- Kanpla API ----------

def noegle():
    return NOEGLE_FIL.read_text(encoding="utf-8").strip()


def eksport(type_id, kolonner, filtre):
    """Kører en eksport og returnerer rækkerne som liste af dicts."""
    krop = {
        "exportFileName": type_id,
        "exportTypeId": type_id,
        "fileFormat": "json",
        "columns": [{"type": "read", "name": k} for k in kolonner],
        "filters": [{"column": c, "operator": o, "value": v} for c, o, v in filtre],
    }
    req = urllib.request.Request(EXPORT_URL, data=json.dumps(krop).encode(), method="POST",
                                 headers={"x-api-key": noegle(), "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=180) as r:
        url = json.load(r)["url"]
    with urllib.request.urlopen(url, timeout=180) as r:     # signeret URL, kortlivet
        return json.load(r)


def salgslinjer(fra, til, lokation_id=None):
    """Én række pr. solgt produkt, fra og med `fra` til og med `til`."""
    filtre = [("date", ">=", fra.isoformat()), ("date", "<", (til + timedelta(days=1)).isoformat())]
    if lokation_id:
        filtre.append(("school_id", "=", lokation_id))
    return eksport("saleItems",
                   ["date", "school_name", "module_name", "product_name", "product_amount", "order_or_refund"],
                   filtre)


def gaester_pr_dag(fra, til, lokation=None, lokation_id=None):
    """{(lokation, dato): antal}. Refunderinger har negativt antal i Kanpla og trækker derfor selv fra."""
    ud = defaultdict(int)
    for r in salgslinjer(fra, til, lokation_id):
        sted = (r["school_name"] or "").strip()
        if lokation and sted.lower() != lokation.lower():
            continue
        if GAESTE_MODUL not in (r["module_name"] or "").lower():
            continue
        ud[(sted, r["date"])] += r["product_amount"] or 0
    return dict(ud)


# ---------- Excel ----------

def opsaet(wb):
    """Opretter fanen Kanpla (tabel), kolonnen "Kuvert manuel" og Kuverter-formlerne i månedsfanerne."""
    navne = [ws.Name for ws in wb.Worksheets]
    if FANE not in navne:
        ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))   # efter sidste fane
        ws.Name = FANE
        ws.Range(ws.Cells(1, 1), ws.Cells(1, len(KOLONNER))).Value = (tuple(KOLONNER),)
        lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(1, 1), ws.Cells(2, len(KOLONNER))), None, 1)
        lo.Name = FANE
        lo.TableStyle = "TableStyleMedium2"
        for i, b in enumerate([12, 11, 6, 16, 9, 17], 1):
            ws.Columns(i).ColumnWidth = b
        intl = wb.Application.International          # formaterne skrives i Excels egne (danske) tegn
        d, m, y = intl[20], intl[19], intl[18]
        ws.Columns(1).NumberFormatLocal = f"{d * 2}-{m * 2}-{y * 4}"
        ws.Activate()
        wb.Windows(1).SplitRow = 1
        wb.Windows(1).FreezePanes = True

    skabelon = next((wb.Worksheets(md) for md in MDR if md in navne
                     and wb.Worksheets(md).Range("O8").Value == "Kuvert manuel"), None)
    for md in MDR:
        if md not in navne:
            continue
        ws = wb.Worksheets(md)
        if ws.Range("O8").Value != "Kuvert manuel":
            # Samme opsætning som den måned, der allerede har kolonnen: indsæt "Kuvert manuel" foran Kuverter
            assert ws.Range("O8").Value == "Kuverter", f"{md}: uventet overskrift i O8: {ws.Range('O8').Value}"
            ws.Columns("O").Insert()
            if skabelon is not None:                         # formatering og overskrift som skabelonen
                skabelon.Range("O1:O41").Copy(ws.Range("O1:O41"))
                ws.Columns("O").ColumnWidth = skabelon.Columns("O").ColumnWidth
                ws.Range("O9:O39").ClearContents()           # skabelonens manuelle tal skal ikke med
            else:
                ws.Range("O8").Value = "Kuvert manuel"
        for r in range(9, 40):
            sumifs = f"SUMIFS({FANE}[Gæster],{FANE}[Måned],$B$2,{FANE}[Dag],$A{r})"
            ws.Range(f"P{r}").Formula = f'=IF($A{r}="",0,IF($O{r}<>"",$O{r},{sumifs}))'
        ws.Range("P9:P39").NumberFormatLocal = "0;-0;"     # 0 vises som tom celle
    wb.Application.CutCopyMode = False
    wb.Worksheets(1).Activate()


def opdater_ark(wb, fra=None, til=None):
    """Henter gæster for LOKATION og skriver/overskriver dem pr. dato i Kanpla-fanen. Returnerer antal ændrede dage."""
    til = til or date.today()
    fra = fra or max(date(til.year, 1, 1), til - timedelta(days=DAGE_TILBAGE))
    lo = wb.Worksheets(FANE).ListObjects(FANE)
    body = lo.DataBodyRange
    eksisterende = {}                          # dato -> (række, gæster)
    if body is not None:
        rows = body.Value
        for v in rows:
            if v[0] not in (None, ""):
                d = v[0].date() if hasattr(v[0], "date") else date.fromordinal(date(1899, 12, 30).toordinal() + int(v[0]))
                eksisterende[d] = v
    if not eksisterende:                       # første gang: hent hele året
        fra = date(til.year, 1, 1)

    nye = {date.fromisoformat(d): n for (_, d), n in gaester_pr_dag(fra, til, LOKATION, LOKATION_ID).items()}
    hentet = datetime.now().strftime("%Y-%m-%d %H:%M")
    aendret = 0
    for d, n in nye.items():
        gammel = eksisterende.get(d)
        if gammel is None or int(gammel[4] or 0) != n:
            eksisterende[d] = (d, MDR[d.month - 1], d.day, LOKATION, n, hentet)
            aendret += 1
    if not aendret:
        return 0

    def serie(d):
        return (d - date(1899, 12, 30)).days
    raekker = [(serie(d),) + tuple(v[1:]) for d, v in sorted(eksisterende.items())]
    ws = lo.Parent
    start = lo.HeaderRowRange.Row + 1
    lo.Resize(ws.Range(lo.HeaderRowRange.Cells(1, 1), ws.Cells(start + len(raekker) - 1, len(KOLONNER))))
    ws.Range(ws.Cells(start, 1), ws.Cells(start + len(raekker) - 1, len(KOLONNER))).Value = raekker
    return aendret


# ---------- Kommandolinje ----------

def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("fra", nargs="?", type=date.fromisoformat)
    p.add_argument("til", nargs="?", type=date.fromisoformat)
    p.add_argument("--lokation")
    p.add_argument("--moduler", action="store_true", help="vis antal pr. modul i stedet for gæster pr. dag")
    p.add_argument("--opsaet", action="store_true", help="opret Kanpla-fanen og Kuverter-formlerne i arket")
    p.add_argument("--opdater", action="store_true", help="hent gæster ind i arket nu")
    a = p.parse_args()

    if a.opsaet or a.opdater:
        from ordre_til_ark import ARK, aabn_ark
        xl, wb, egen = aabn_ark(ARK)
        try:
            if a.opsaet:
                opsaet(wb)
            n = opdater_ark(wb, a.fra, a.til)
            wb.Save()
            print(f"Kanpla: {n} dage opdateret for {LOKATION}")
        finally:
            if egen:
                wb.Close(False)
                xl.Quit()
        return

    if not (a.fra and a.til):
        p.error("angiv fra- og til-dato, eller --opsaet / --opdater")
    if a.moduler:
        pr_modul = defaultdict(int)
        for r in salgslinjer(a.fra, a.til):
            if a.lokation and (r["school_name"] or "").strip().lower() != a.lokation.lower():
                continue
            pr_modul[((r["school_name"] or "").strip(), (r["module_name"] or "").strip())] += r["product_amount"] or 0
        for (sted, modul), n in sorted(pr_modul.items()):
            print(f"{sted:35} {modul:55} {n:6}")
        return

    for (sted, dag), n in sorted(gaester_pr_dag(a.fra, a.til, a.lokation).items()):
        print(f"{sted:35} {dag}  {n:5}")


if __name__ == "__main__":
    main()
