"""Læser en "Ordrekvittering"-mail (fx fra AB Catering København) og returnerer ordren som en dict.

Understøtter:
  * .eml  – en gemt mail (som Outlook/Power Automate kan gemme den)
  * .html – mailens HTML-krop
  * .txt  – mailens tekst-krop (fallback, når der ikke er HTML)
"""
import email
import re
from datetime import datetime
from email import policy
from html.parser import HTMLParser
from pathlib import Path

HOVED_FELTER = {
    "ordre": "ordre",
    "leveringsdato": "leveringsdato",
    "kunde": "kunde",
    "bruger": "bruger",
    "leveringsbemærkning": "leveringsbemaerkning",
}


def tal(s):
    """'1.234,56' -> 1234.56"""
    s = s.strip().replace(".", "").replace(" ", "").replace(",", ".")
    return float(s) if s else 0.0


class _TabelParser(HTMLParser):
    """Samler alle tabelrækker som lister af celletekster, plus al tekst i rækkefølge."""

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


def _laes_fil(sti):
    sti = Path(sti)
    if sti.suffix.lower() == ".eml":
        msg = email.message_from_bytes(sti.read_bytes(), policy=policy.default)
        afsender = str(msg.get("From", ""))
        del_ = msg.get_body(preferencelist=("html", "plain"))
        indhold = del_.get_content() if del_ else ""
        return indhold, del_ is not None and del_.get_content_subtype() == "html", afsender
    tekst = sti.read_text(encoding="utf-8")
    return tekst, sti.suffix.lower() in (".html", ".htm"), ""


def _fra_tabeller(raekker, ordre):
    for r in raekker:
        celler = [c for c in r if c]
        if not celler:
            continue
        noegle = celler[0].rstrip(":").strip().lower()
        if noegle in HOVED_FELTER and r[0].strip().endswith(":"):
            ordre[HOVED_FELTER[noegle]] = celler[1] if len(celler) > 1 else ""
        elif re.fullmatch(r"\d{4,}", celler[0]) and len(celler) >= 6:
            varenr, vare, antal, enhed, enhedspris, linjepris = celler[:6]
            ordre["linjer"].append({
                "varenr": varenr, "vare": vare, "antal": tal(antal), "enhed": enhed,
                "enhedspris": tal(enhedspris), "linjepris": tal(linjepris),
            })
        elif celler[0].lower() == "total" and len(celler) >= 2:
            ordre["total"] = tal(celler[-1])


_LINJE_RE = re.compile(
    r"^(?P<varenr>\d{4,})\s+(?P<vare>.+?)\s+(?P<antal>[\d.,]+)\s+(?P<enhed>[A-ZÆØÅa-zæøå]+)\s+"
    r"(?P<enhedspris>[\d.]*\d,\d{2})\s+(?P<linjepris>[\d.]*\d,\d{2})\s*$"
)


def _fra_tekst(tekst, ordre):
    """Fallback for ren tekst: én varelinje pr. linje."""
    for linje in (l.strip() for l in tekst.splitlines()):
        m = re.match(r"^(Ordre|Leveringsdato|Kunde|Bruger|Leveringsbemærkning):\s*(.*)$", linje, re.I)
        if m:
            ordre[HOVED_FELTER[m.group(1).lower()]] = m.group(2).strip()
            continue
        m = _LINJE_RE.match(linje)
        if m:
            d = m.groupdict()
            ordre["linjer"].append({
                "varenr": d["varenr"], "vare": d["vare"], "antal": tal(d["antal"]), "enhed": d["enhed"],
                "enhedspris": tal(d["enhedspris"]), "linjepris": tal(d["linjepris"]),
            })
            continue
        m = re.match(r"^Total\s+([\d.]*\d,\d{2})$", linje, re.I)
        if m:
            ordre["total"] = tal(m.group(1))


def find_regel(tekst, kunde, regler):
    """Første kolonneregel hvor leverandøren genkendes i mailen og (hvis udfyldt) kunden indeholder teksten.
    Samme logik som findRegel() i office-scripts/IndlaesOrdrekvittering.ts."""
    for r in regler:
        if (r.get("genkend") or r["leverandoer"]).lower() in tekst.lower() and                 (not r.get("kunde_indeholder") or r["kunde_indeholder"].lower() in kunde.lower()):
            return r
    return None


def parse_ordrekvittering(sti, regler):
    indhold, er_html, afsender = _laes_fil(sti)
    ordre = {"ordre": "", "leveringsdato": "", "kunde": "", "bruger": "", "leveringsbemaerkning": "",
             "linjer": [], "total": None, "kilde": Path(sti).name}

    if er_html:
        p = _TabelParser()
        p.feed(indhold)
        tekst = "".join(p.tekst)
        _fra_tabeller(p.raekker, ordre)
    else:
        tekst = indhold
    if not ordre["linjer"]:
        _fra_tekst(tekst, ordre)

    regel = find_regel(tekst + " " + afsender, ordre["kunde"], regler)
    ordre["leverandoer"] = regel["leverandoer"] if regel else "Ukendt"
    ordre["kolonne"] = regel["kolonne"] if regel else ""

    if ordre["leveringsdato"]:
        ordre["leveringsdato"] = datetime.strptime(ordre["leveringsdato"], "%d/%m/%Y").date().isoformat()

    sum_linjer = round(sum(l["linjepris"] for l in ordre["linjer"]), 2)
    if ordre["total"] is None:
        ordre["total"] = sum_linjer
    ordre["advarsler"] = []
    if not ordre["ordre"]:
        ordre["advarsler"].append("Ordrenummer ikke fundet")
    if not ordre["linjer"]:
        ordre["advarsler"].append("Ingen varelinjer fundet")
    if abs(sum_linjer - ordre["total"]) > 0.05:
        ordre["advarsler"].append(f"Sum af linjer ({sum_linjer}) matcher ikke total ({ordre['total']})")
    if not regel:
        ordre["advarsler"].append("Ukendt leverandør – tilføj den i Kolonneregler")
    return ordre

