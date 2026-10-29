"""Indlæser Kanpla-eksport (kuverter + ekstrasalg pr. dag).

Pilot-format (semikolon-separeret, beløb inkl. moms, dansk decimalkomma):
  dato;kuverter;moedeforplejning;fredagsmorgenmad_antal;smart_fridge;kaffeordning;frugtordning;special_arrangementer;boden_takeaway
Når vi har en rigtig Kanpla-eksport, tilpasses kolonnenavnene her.
"""
import csv
from pathlib import Path


def _tal(s):
    s = (s or "").strip().replace(".", "").replace(",", ".") if "," in (s or "") else (s or "").strip()
    return float(s) if s else 0.0


def laes_kanpla(mappe):
    dage = {}
    for fil in sorted(Path(mappe).glob("*.csv")):
        with open(fil, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f, delimiter=";"):
                dage[r["dato"]] = {k: _tal(v) for k, v in r.items() if k != "dato"}
    return dage
