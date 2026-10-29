"""Laver DEMO-data til piloten: ordrekvitterings-mails fra AB Catering + en Kanpla-eksport for oktober 2026.

Ordre 31525653 er gengivet præcis som den rigtige mail. Resten er opdigtede, så dashboardet har noget at vise.
Kør:  python scripts/lav_demodata.py
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

ROD = Path(__file__).resolve().parent.parent
INDBAKKE = ROD / "eksempler" / "indbakke"
KANPLA = ROD / "eksempler" / "kanpla"

SKABELON = """<html><body style="background:#2b2b2b;color:#fff;font-family:Arial">
<p style="text-align:right">Ordrekvittering</p>
<h2>Tak for din bestilling</h2>
<p>Vi har modtaget din bestilling og er i gang med at behandle den. Bemærk venligst, at dette ikke er en
ordrebekræftelse, men alene en kopi af din ordre.</p>
<table>
<tr><td>Ordre:</td><td>{ordre}</td></tr>
<tr><td>Leveringsdato:</td><td>{dato}</td></tr>
<tr><td>Kunde:</td><td>{kunde}</td></tr>
<tr><td>Bruger:</td><td><a href="mailto:kaso@kokkeneskoekken.dk">kaso@kokkeneskoekken.dk</a></td></tr>
<tr><td>Leveringsbemærkning:</td><td>{bem}</td></tr>
</table>
<table>
<tr><th>VARENR.</th><th>VARE</th><th>ANTAL</th><th>ENHED</th><th>ENHEDSPRIS</th><th>LINJEPRIS</th></tr>
{linjer}
<tr><td></td><td></td><td></td><td></td><td>Total</td><td>{total}</td></tr>
</table>
<p>Med venlig hilsen<br>AB Catering København</p>
</body></html>
"""


def kr(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def lav_mail(ordre, dato, kunde, varer, bem=""):
    linjer, total = [], 0.0
    for varenr, navn, antal, enhed, pris in varer:
        lp = round(antal * pris, 2)
        total += lp
        linjer.append(f"<tr><td>{varenr}</td><td>{navn}</td><td>{antal}</td><td>{enhed}</td>"
                      f"<td>{kr(pris)}</td><td>{kr(lp)}</td></tr>")
    html = SKABELON.format(ordre=ordre, dato=dato.strftime("%d/%m/%Y"), kunde=kunde, bem=bem,
                           linjer="\n".join(linjer), total=kr(total))
    (INDBAKKE / f"ordrekvittering_{ordre}.html").write_text(html, encoding="utf-8")


FRUGT = [
    ("224918", "Æble Ingrid Marie DK Øko", "STK", 3.42),
    ("306708", "Banan Gul STK Øko", "STK", 2.85),
    ("206165", "Blommer Øko", "PS", 29.39),
    ("380219", "Havesyre Rød i PS Øko", "PS", 20.40),
    ("390612", "Tatsoi Mini Øko", "KRT", 161.25),
    ("201144", "Pære Conference Øko", "STK", 3.95),
    ("201877", "Clementin Øko", "KG", 32.50),
]
CATERING = [
    ("510233", "Rugbrød Groft Skåret", "STK", 28.50),
    ("520410", "Kyllingebryst Fersk DK", "KG", 89.00),
    ("530118", "Gulerødder Øko 10 kg", "SK", 96.00),
    ("540027", "Kartofler Nye DK 15 kg", "SK", 145.00),
    ("550301", "Mælk Let Øko 1 L", "STK", 13.75),
    ("560911", "Ost Skiver 45+ 1 kg", "KG", 78.00),
    ("570602", "Salatmix Øko 1 kg", "PS", 64.00),
    ("580044", "Hytteost 2 kg", "STK", 69.00),
]


def main():
    INDBAKKE.mkdir(parents=True, exist_ok=True)
    KANPLA.mkdir(parents=True, exist_ok=True)
    random.seed(42)

    # Den rigtige mail fra 27/10/2026
    lav_mail(31525653, date(2026, 10, 27), "39-69505002, FORENEDE SERVICE FRUGT", [
        ("224918", "Æble Ingrid Marie DK Øko", 20, "STK", 3.42),
        ("306708", "Banan Gul STK Øko", 20, "STK", 2.85),
        ("206165", "Blommer Øko", 2, "PS", 29.39),
        ("380219", "Havesyre Rød i PS Øko", 2, "PS", 20.40),
        ("390612", "Tatsoi Mini Øko", 1, "KRT", 161.25),
    ])

    # Opdigtede ordrer resten af oktober
    ordrenr = 31525000
    d = date(2026, 10, 1)
    rows = []
    while d.month == 10:
        hverdag = d.weekday() < 5
        if hverdag and d != date(2026, 10, 27):
            ordrenr += 7
            lav_mail(ordrenr, d, "39-69505002, FORENEDE SERVICE FRUGT",
                     [(n, v, random.choice([5, 10, 20, 30]) if e == "STK" else random.randint(1, 3), e, p)
                      for n, v, e, p in random.sample(FRUGT, 4)])
            ordrenr += 7
            lav_mail(ordrenr, d, "39-69505001, FORENEDE SERVICE KANTINE",
                     [(n, v, random.randint(4, 25) if e == "STK" else random.randint(2, 12), e, p)
                      for n, v, e, p in random.sample(CATERING, 5)])
        kuv = random.randint(85, 115) if hverdag else 0
        rows.append({
            "dato": d.isoformat(),
            "kuverter": kuv,
            "moedeforplejning": round(random.choice([0, 100, 260, 648.75, 1335, 2197.5]), 2) if hverdag else 0,
            "fredagsmorgenmad_antal": 100 if d.weekday() == 4 else 0,
            "smart_fridge": round(random.uniform(0, 300), 2) if hverdag else 0,
            "kaffeordning": round(random.choice([0, 0, 308.9, 459.3, 828.35]), 2) if hverdag else 0,
            "frugtordning": round(random.uniform(100, 600), 2) if hverdag else 0,
            "special_arrangementer": 0,
            "boden_takeaway": round(random.uniform(0, 250), 2) if hverdag else 0,
        })
        d += timedelta(days=1)

    with open(KANPLA / "kanpla_2026-10.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]), delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow({k: (str(v).replace(".", ",") if isinstance(v, float) else v) for k, v in r.items()})
    print(f"Demo-mails i {INDBAKKE}, Kanpla-fil i {KANPLA}")


if __name__ == "__main__":
    main()
