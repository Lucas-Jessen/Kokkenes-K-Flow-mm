"""Henter ordrekvitteringer fra kundeservice@ab-kbh.dk i indkoebkk@forenede.dk (Outlook) og skriver varelinjerne ind i fanen "Ordreliste" i Kokkenes k ark.xlsx.

Månedsfanerne henter selv beløbene fra Ordreliste med formler (pr. leveringsdato og kolonne).
Ordrenummeret bruges som kontrol: en ordre, der allerede står i Ordreliste, skrives ikke igen.
Indlæste mails markeres som læst. Mails, der ikke kan læses som en ordre, forbliver ulæste.

Brug:
  python ordre_til_ark.py                    # læs postkassen og opdatér arket
  python ordre_til_ark.py --mappe <mappe>    # læs gemte mails (.msg/.eml/.html) i stedet for postkassen
  python ordre_til_ark.py --opsaet           # opret fanen Ordreliste + formler i månedsfanerne (kun én gang)
  python ordre_til_ark.py --vis-postkasser   # vis hvilke postkasser Outlook kan se

Mails hentes via Microsoft Graph, hvis graph.json findes (se graph_mail.py) – ellers via klassisk Outlook.
Kræver Excel og pakken pywin32.
"""
import argparse
import email
import re
import sys
from datetime import datetime, timedelta
from email import policy
from html.parser import HTMLParser
from pathlib import Path

import win32com.client

import graph_mail
import kanpla
import rapport

ROD = Path(__file__).resolve().parent
ARK = ROD / "Kokkenes k ark.xlsx"
POSTKASSE = "indkoebkk@forenede.dk"
AFSENDER = "kundeservice@ab-kbh.dk"     # kun mails fra (eller videresendt fra) denne afsender behandles
DAGE_TILBAGE = 60            # hvor langt tilbage i indbakken der kigges ved hver kørsel

# Hvilken kolonne i månedsfanerne en ordre tæller i. Første regel, der passer, vinder.
# "kolonne" skal stå præcis som overskriften i række 8.
# Pilot: alt fra AB Catering tæller under "BC catering", uanset kunde.
KOLONNEREGLER = [
    {"leverandoer": "AB Catering København", "genkend": "AB Catering", "kunde_indeholder": "", "kolonne": "BC catering"},
]

MDR = ["Januar", "Februar", "Marts", "April", "Maj", "Juni", "Juli", "August",
       "September", "Oktober", "November", "December"]
KOLONNER = ["Ordre", "Leveringsdato", "Måned", "Dag", "Leverandør", "Kolonne", "Kunde", "Bruger",
            "Leveringsbemærkning", "Varenr", "Vare", "Antal", "Enhed", "Enhedspris", "Linjepris",
            "Ordre total", "Mail modtaget", "Advarsler"]
FANE = "Ordreliste"


# ---------- Læs mailen ----------

def tal(s):
    """'1.234,56' -> 1234.56. Tekst omkring tallet ignoreres ('Kg. pris 214,97' -> 214.97)."""
    m = re.search(r"-?\d[\d.]*(?:,\d+)?", s.replace(" ", ""))
    return float(m[0].replace(".", "").replace(",", ".")) if m else 0.0


class _TabelParser(HTMLParser):
    """Samler alle tabelrækker som lister af celletekster, plus al tekst."""

    def __init__(self):
        super().__init__()
        self.raekker, self.tekst = [], []
        self._raekke, self._celle = None, None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._raekke = []
        elif tag in ("td", "th"):
            self._celle = []
        elif tag in ("br", "p", "div"):
            self.tekst.append("\n")

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._celle is not None:
            if self._raekke is not None:
                self._raekke.append(" ".join("".join(self._celle).split()))
            self._celle = None
            self.tekst.append("\n")
        elif tag == "tr" and self._raekke is not None:
            if any(self._raekke):
                self.raekker.append(self._raekke)
            self._raekke = None

    def handle_data(self, data):
        if self._celle is not None:
            self._celle.append(data)
        self.tekst.append(data)


HOVED = {"ordre": "ordre", "leveringsdato": "leveringsdato", "kunde": "kunde", "bruger": "bruger",
         "leveringsbemærkning": "leveringsbemaerkning"}
LINJE_RE = re.compile(r"^(\d{4,})\s+(.+?)\s+([\d.,]+)\s+([A-ZÆØÅa-zæøå]+)\s+([\d.]*\d,\d{2})\s+([\d.]*\d,\d{2})\s*$")


def parse_ordrekvittering(html_eller_tekst):
    o = {"ordre": "", "leveringsdato": "", "kunde": "", "bruger": "", "leveringsbemaerkning": "",
         "linjer": [], "total": None}
    er_html = re.search(r"<\s*(td|tr|table|br|p|div)\b", html_eller_tekst, re.I) is not None
    if er_html:
        p = _TabelParser()
        p.feed(html_eller_tekst)
        tekst = "".join(p.tekst)
        for r in p.raekker:
            celler = [c for c in r if c]
            if not celler:
                continue
            noegle = celler[0].rstrip(":").strip().lower()
            if noegle in HOVED and r[0].strip().endswith(":"):
                o[HOVED[noegle]] = celler[1] if len(celler) > 1 else ""
            elif re.fullmatch(r"\d{4,}", celler[0]) and len(celler) >= 6:
                # Vægtvarer har "Kg. pris 214,97" som enhedspris; linjeprisen er den faktiske pris
                enhed = celler[3] + (" (kg-pris)" if re.search(r"kg", celler[4], re.I) else "")
                o["linjer"].append(dict(zip(("varenr", "vare", "antal", "enhed", "enhedspris", "linjepris"),
                                            celler[:2] + [tal(celler[2]), enhed, tal(celler[4]), tal(celler[5])])))
            elif celler[0].lower() == "total" and len(celler) >= 2:
                o["total"] = tal(celler[-1])
    else:
        tekst = html_eller_tekst
    if not o["linjer"]:                 # fallback: ren tekst, én varelinje pr. linje
        for linje in (l.strip() for l in tekst.splitlines()):
            m = re.match(r"^(Ordre|Leveringsdato|Kunde|Bruger|Leveringsbemærkning):\s*(.*)$", linje, re.I)
            if m:
                o[HOVED[m.group(1).lower()]] = m.group(2).strip()
            elif (m := LINJE_RE.match(linje)):
                o["linjer"].append({"varenr": m[1], "vare": m[2], "antal": tal(m[3]), "enhed": m[4],
                                    "enhedspris": tal(m[5]), "linjepris": tal(m[6])})
            elif (m := re.match(r"^Total\s+([\d.]*\d,\d{2})$", linje, re.I)):
                o["total"] = tal(m[1])

    regel = next((r for r in KOLONNEREGLER
                  if r["genkend"].lower() in tekst.lower()
                  and (not r["kunde_indeholder"] or r["kunde_indeholder"].lower() in o["kunde"].lower())), None)
    o["leverandoer"] = regel["leverandoer"] if regel else "Ukendt"
    o["kolonne"] = regel["kolonne"] if regel else ""

    sum_linjer = round(sum(l["linjepris"] for l in o["linjer"]), 2)
    if o["total"] is None:
        o["total"] = sum_linjer
    adv = []
    if not regel:
        adv.append("Ukendt leverandør")
    if abs(sum_linjer - o["total"]) > 0.05:
        adv.append(f"Sum af linjer {sum_linjer} ≠ total {o['total']}")
    try:
        o["dato"] = datetime.strptime(o["leveringsdato"], "%d/%m/%Y")
    except ValueError:
        o["dato"] = None
        adv.append(f"Leveringsdato kan ikke læses: '{o['leveringsdato']}'")
    o["advarsler"] = "; ".join(adv)
    return o


# ---------- Hent mails ----------

def find_indbakke(ns):
    """Postkassen som ekstra konto i Outlook, ellers som delt postkasse."""
    def norm(s):                                     # "Indkøb KK" -> "indkoebkk"
        s = s.lower().replace("æ", "ae").replace("ø", "oe").replace("å", "aa")
        return re.sub(r"[^a-z0-9]", "", s)
    for st in ns.Stores:
        if norm(POSTKASSE.split("@")[0]) == norm(st.DisplayName):
            try:
                return st.GetDefaultFolder(6)        # 6 = Indbakke
            except Exception:
                continue                             # Outlook kan vise samme postkasse to gange
    r = ns.CreateRecipient(POSTKASSE)
    r.Resolve()
    return ns.GetSharedDefaultFolder(r, 6)


def mails_fra_outlook():
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    try:
        indbakke = find_indbakke(ns)
    except Exception as e:
        sys.exit(f"Kan ikke åbne postkassen {POSTKASSE} i Outlook ({e}).\n"
                 f"Tilføj postkassen i Outlook (Filer › Kontoindstillinger), og kør igen.\n"
                 f"Postkasser Outlook kan se: {', '.join(s.DisplayName for s in ns.Stores)}")
    fra = (datetime.now() - timedelta(days=DAGE_TILBAGE)).strftime("%d-%m-%Y %H:%M")
    items = indbakke.Items.Restrict(f"[ReceivedTime] >= '{fra}'")
    for m in list(items):
        # Direkte fra AB – eller videresendt, hvor AB's adresse står i mailen ("Fra: kundeservice@ab-kbh.dk")
        if getattr(m, "Class", None) == 43 and (AFSENDER.lower() in afsender(m).lower()
                                                 or AFSENDER.lower() in (m.Body or "").lower()):
            yield m.HTMLBody or m.Body, m.ReceivedTime.strftime("%Y-%m-%d %H:%M"), m.Subject, _markoer(m)


def _markoer(m):
    def markér_læst():
        if m.UnRead:
            m.UnRead = False
            m.Save()
    return markér_læst


def afsender(m):
    """SMTP-adressen på afsenderen (også når Exchange gemmer den som intern adresse)."""
    try:
        if m.SenderEmailType == "EX":
            return m.Sender.GetExchangeUser().PrimarySmtpAddress or ""
        return m.SenderEmailAddress or ""
    except Exception:
        return ""


def mails_fra_mappe(mappe):
    ol = None
    for fil in sorted(Path(mappe).iterdir()):
        ext = fil.suffix.lower()
        if ext == ".msg":
            ol = ol or win32com.client.Dispatch("Outlook.Application")
            m = ol.Session.OpenSharedItem(str(fil))
            yield m.HTMLBody or m.Body, m.ReceivedTime.strftime("%Y-%m-%d %H:%M"), fil.name, None
        elif ext == ".eml":
            msg = email.message_from_bytes(fil.read_bytes(), policy=policy.default)
            krop = msg.get_body(preferencelist=("html", "plain"))
            yield (krop.get_content() if krop else ""), str(msg.get("Date", "")), fil.name, None
        elif ext in (".html", ".htm", ".txt"):
            yield fil.read_text(encoding="utf-8"), "", fil.name, None


# ---------- Excel ----------

def aabn_ark(sti):
    """Bruger arket, hvis det allerede er åbent i Excel – ellers åbnes det skjult."""
    sti = str(Path(sti).resolve())
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        for wb in xl.Workbooks:
            if Path(wb.FullName).name.lower() == Path(sti).name.lower():
                return xl, wb, False
    except Exception:
        pass
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible, xl.DisplayAlerts = False, False
    return xl, xl.Workbooks.Open(sti), True


def opsaet(wb):
    """Opretter fanen Ordreliste (tabel) og SUMIFS-formler i månedsfanerne for de kolonner, mails fylder."""
    navne = [ws.Name for ws in wb.Worksheets]
    if FANE not in navne:
        ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))   # efter sidste fane
        ws.Name = FANE
        ws.Range(ws.Cells(1, 1), ws.Cells(1, len(KOLONNER))).Value = (tuple(KOLONNER),)
        lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(1, 1), ws.Cells(2, len(KOLONNER))), None, 1)
        lo.Name = FANE
        lo.TableStyle = "TableStyleMedium2"
        bredder = [11, 13, 10, 6, 22, 13, 36, 26, 22, 9, 30, 8, 8, 12, 12, 12, 17, 30]
        for i, b in enumerate(bredder, 1):
            ws.Columns(i).ColumnWidth = b
        # Formaterne skrives i Excels egne (danske) tegn, ellers vises de forkert
        intl = wb.Application.International          # pywin32 giver hele listen (Excels index - 1)
        d, m, y, dec, tus = intl[20], intl[19], intl[18], intl[2], intl[3]
        k = {n: i + 1 for i, n in enumerate(KOLONNER)}
        ws.Columns(k["Leveringsdato"]).NumberFormatLocal = f"{d * 2}-{m * 2}-{y * 4}"
        for n in ("Enhedspris", "Linjepris", "Ordre total"):
            ws.Columns(k[n]).NumberFormatLocal = f"#{tus}##0{dec}00"
        ws.Activate()
        wb.Windows(1).SplitRow = 1
        wb.Windows(1).FreezePanes = True

    maal = {r["kolonne"] for r in KOLONNEREGLER}
    for md in MDR:
        if md not in navne:
            continue
        ws = wb.Worksheets(md)
        for c in range(2, 9):                                  # B..H = leverandørkolonnerne
            if ws.Cells(8, c).Value not in maal:
                # Kolonnen fyldes ikke af mails: fjern evt. gamle Ordreliste-formler, så den kan tastes manuelt igen
                for r in range(9, 40):
                    if FANE in str(ws.Cells(r, c).Formula):
                        ws.Cells(r, c).ClearContents()
                continue
            bogstav = ws.Cells(8, c).Address.split("$")[1]
            for r in range(9, 40):
                f = (f"SUMIFS({FANE}[Linjepris],{FANE}[Måned],$B$2,{FANE}[Dag],$A{r},"
                     f"{FANE}[Kolonne],{bogstav}$8)")
                ws.Cells(r, c).Formula = f'=IF($A{r}="","",IF({f}=0,"",{f}))'
    wb.Worksheets(1).Activate()


def eksisterende_ordrer(wb):
    lo = wb.Worksheets(FANE).ListObjects(FANE)
    if lo.DataBodyRange is None:
        return set()
    vaerdier = lo.ListColumns("Ordre").DataBodyRange.Value
    if not isinstance(vaerdier, tuple):        # én celle giver en enkelt værdi i stedet for en tabel
        vaerdier = ((vaerdier,),)
    return {str(int(v[0])) if isinstance(v[0], float) else str(v[0]) for v in vaerdier if v[0] not in (None, "")}


def excel_dato(d):
    """Datoen som Excel-serienummer (undgår at Outlook/Excel flytter den en dag pga. tidszone)."""
    return (d - datetime(1899, 12, 30)).days if d else ""


def skriv_ordrer(wb, ordrer):
    lo = wb.Worksheets(FANE).ListObjects(FANE)
    raekker = []
    for o, modtaget in ordrer:
        d = o["dato"]
        for l in o["linjer"]:
            raekker.append((o["ordre"], excel_dato(d), MDR[d.month - 1] if d else "", d.day if d else "", o["leverandoer"],
                            o["kolonne"], o["kunde"], o["bruger"], o["leveringsbemaerkning"], l["varenr"], l["vare"],
                            l["antal"], l["enhed"], l["enhedspris"], l["linjepris"], o["total"], modtaget,
                            o["advarsler"]))
    if not raekker:
        return
    body = lo.DataBodyRange
    # En helt tom tabel har én tom række, som fyldes først
    tom = body is not None and body.Rows.Count == 1 and all(v in (None, "") for v in body.Value[0])
    start = body.Row if tom else (body.Row + body.Rows.Count if body is not None else lo.HeaderRowRange.Row + 1)
    ws = lo.Parent
    lo.Resize(ws.Range(lo.HeaderRowRange.Cells(1, 1), ws.Cells(start + len(raekker) - 1, len(KOLONNER))))
    ws.Range(ws.Cells(start, 1), ws.Cells(start + len(raekker) - 1, len(KOLONNER))).Value = raekker


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--ark", default=ARK, type=Path)
    ap.add_argument("--mappe", type=Path, help="læs gemte mails fra en mappe i stedet for Outlook")
    ap.add_argument("--opsaet", action="store_true", help="opret Ordreliste-fanen og formlerne")
    ap.add_argument("--vis-postkasser", action="store_true")
    a = ap.parse_args()

    if a.vis_postkasser:
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        print("\n".join(s.DisplayName for s in ns.Stores))
        return

    xl, wb, egen = aabn_ark(a.ark)
    try:
        if a.opsaet:
            opsaet(wb)
            wb.Save()
            print(f"Fanen '{FANE}' og formlerne er sat op i {a.ark.name}")
            return
        if FANE not in [ws.Name for ws in wb.Worksheets]:
            sys.exit(f"Fanen '{FANE}' findes ikke – kør først med --opsaet")

        kendte = eksisterende_ordrer(wb)
        nye, dubletter, fejl = [], 0, 0
        laest = []                                   # mails der markeres som læst, når arket er gemt
        if a.mappe:
            kilde = mails_fra_mappe(a.mappe)
        elif graph_mail.konfigureret():              # Microsoft Graph: ingen Outlook, ingen pop-ups
            kilde = graph_mail.mails(POSTKASSE, AFSENDER, DAGE_TILBAGE)
        else:
            kilde = mails_fra_outlook()
        for html, modtaget, navn, markér_læst in kilde:
            try:
                o = parse_ordrekvittering(html)
            except Exception as e:                   # én mærkelig mail må ikke stoppe resten
                fejl += 1
                print(f"  SPRUNGET OVER: {navn} – kunne ikke læses ({e})")
                continue
            if not o["ordre"] or not o["linjer"]:
                fejl += 1                            # forbliver ulæst, så nogen kigger på den
                print(f"  SPRUNGET OVER: {navn} – ingen ordrenummer/varelinjer fundet")
                continue
            if markér_læst is not None:
                laest.append(markér_læst)
            if o["ordre"] in kendte:
                dubletter += 1
                continue
            kendte.add(o["ordre"])
            nye.append((o, modtaget))
            print(f"  NY: ordre {o['ordre']}  {o['leveringsdato']}  {o['kolonne'] or '?'}  {o['total']:.2f} kr."
                  + (f"  ADVARSEL: {o['advarsler']}" if o["advarsler"] else ""))
        skriv_ordrer(wb, nye)
        if nye:
            wb.Save()
        for markér in laest:
            markér()
        print(f"{len(nye)} nye ordrer, {dubletter} fandtes allerede, {fejl} kunne ikke læses")

        # Spisende gæster fra Kanpla – en fejl her må ikke stoppe ordrerne
        if not a.mappe and kanpla.FANE in [ws.Name for ws in wb.Worksheets]:
            try:
                n = kanpla.opdater_ark(wb)
                if n:
                    wb.Save()
                print(f"Kanpla: {n} dage opdateret for {kanpla.LOKATION}")
            except Exception as e:
                print(f"Kanpla FEJL: {e}")

        if not a.mappe:
            try:
                rapport.skriv_html(rapport.hent_data(wb))
            except Exception as e:
                print(f"Rapport FEJL: {e}")
    finally:
        if egen:
            wb.Close(False)
            xl.Quit()


if __name__ == "__main__":
    if sys.stdout is None:                   # startet skjult (pythonw / Opgavestyring): skriv til logfil
        sys.stdout = sys.stderr = open(ROD / "ordre_til_ark.log", "a", encoding="utf-8", buffering=1)
        print(f"--- {datetime.now():%Y-%m-%d %H:%M}")
        try:
            main()
        except SystemExit as e:
            if e.code not in (None, 0):
                print(e.code)
        except Exception:
            import traceback
            traceback.print_exc()
    else:
        main()
