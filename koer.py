"""K-flow: læs nye ordrekvitteringer + Kanpla-data -> Excel + data til hjemmesiden.

Brug:
  python koer.py                                   # bruger eksempler/indbakke og eksempler/kanpla
  python koer.py --indbakke <mappe> --kanpla <mappe>
"""
import argparse
import json
from datetime import datetime
from pathlib import Path

from kkflow.beregning import beregn_maaned
from kkflow.excel_export import byg_excel
from kkflow.kanpla import laes_kanpla
from kkflow.parse_ordre import parse_ordrekvittering

ROD = Path(__file__).resolve().parent


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--indbakke", default=ROD / "eksempler" / "indbakke", type=Path)
    ap.add_argument("--kanpla", default=ROD / "eksempler" / "kanpla", type=Path)
    ap.add_argument("--data", default=ROD / "data", type=Path)
    ap.add_argument("--output", default=ROD / "output", type=Path)
    a = ap.parse_args()

    cfg = json.loads((ROD / "config.json").read_text(encoding="utf-8"))
    a.data.mkdir(exist_ok=True)
    a.output.mkdir(exist_ok=True)

    # 1) Ordrer – gemmes pr. ordrenummer, så samme mail ikke tælles to gange
    lager = a.data / "ordrer.json"
    ordrer = json.loads(lager.read_text(encoding="utf-8")) if lager.exists() else {}
    nye = 0
    for fil in sorted(p for p in a.indbakke.iterdir() if p.suffix.lower() in (".eml", ".html", ".htm", ".txt")):
        o = parse_ordrekvittering(fil, cfg["leverandoerer"])
        for adv in o["advarsler"]:
            print(f"  ADVARSEL {fil.name}: {adv}")
        if not o["ordre"]:
            continue
        nye += o["ordre"] not in ordrer
        ordrer[o["ordre"]] = o
    lager.write_text(json.dumps(ordrer, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(ordrer)} ordrer i alt ({nye} nye)")

    # 2) Kanpla
    kanpla = laes_kanpla(a.kanpla)
    print(f"{len(kanpla)} dage fra Kanpla")

    # 3) Excel + dashboard-data
    liste = list(ordrer.values())
    maaneder = {tuple(map(int, d[:7].split("-"))) for d in
                [o["leveringsdato"] for o in liste if o["leveringsdato"]] + list(kanpla)}
    xlsx = a.output / "K-flow.xlsx"
    byg_excel(xlsx, liste, kanpla, cfg, maaneder)
    print(f"Excel: {xlsx}")

    web = {
        "opdateret": datetime.now().isoformat(timespec="minutes"),
        "kantine": cfg["kantine"],
        "kuvertpris": cfg["kuvertpris"],
        "kolonner": [k["navn"] for k in cfg["indkoebs_kolonner"]],
        "maaneder": {f"{y}-{m:02d}": beregn_maaned(y, m, liste, kanpla, cfg) for y, m in sorted(maaneder)},
        "ordrer": sorted(liste, key=lambda o: (o["leveringsdato"], str(o["ordre"])), reverse=True),
    }
    js = ROD / "web" / "data.js"
    js.write_text("window.KK_DATA = " + json.dumps(web, ensure_ascii=False) + ";\n", encoding="utf-8")
    print(f"Hjemmeside-data: {js}")


if __name__ == "__main__":
    main()
