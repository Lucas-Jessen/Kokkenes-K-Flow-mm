"""Samme regnelogik som det manuelle månedsark – bruges til dashboardet (Excel-arket får formler)."""
import calendar
from datetime import date


def ekstrasalg_dag(kanpla_dag, cfg):
    """Returnerer (ekstrasalg inkl. moms, vareforbrug ex moms) for én dag."""
    inkl, vare = 0.0, 0.0
    for kat in cfg["ekstrasalg"]:
        v = kanpla_dag.get(kat["noegle"], 0.0) * kat.get("pris_pr_stk", 1)
        inkl += v
        vare += v * kat["sats"]
    return inkl, vare / cfg["moms_faktor"]


def beregn_maaned(aar, md, ordrer, kanpla, cfg):
    kolonner = [k["navn"] for k in cfg["indkoebs_kolonner"]]
    kp = cfg["kuvertpris"]
    dage = []
    mtd = kuv_cum = budget_cum = vare_cum = 0.0
    for dag in range(1, calendar.monthrange(aar, md)[1] + 1):
        iso = date(aar, md, dag).isoformat()
        indkoeb = {k: 0.0 for k in kolonner}
        for o in ordrer:
            if o["leveringsdato"] == iso and o.get("kolonne") in indkoeb:
                indkoeb[o["kolonne"]] += o["total"]
        i_alt = sum(indkoeb.values())
        mtd += i_alt
        k = kanpla.get(iso, {})
        kuverter = k.get("kuverter", 0.0)
        ekstra_inkl, vareforbrug = ekstrasalg_dag(k, cfg)
        kuv_cum += kuverter
        budget_cum += kuverter * kp
        vare_cum += vareforbrug
        dage.append({
            "dato": iso,
            "indkoeb": {kk: round(v, 2) for kk, v in indkoeb.items()},
            "indkoeb_i_alt": round(i_alt, 2),
            "mtd": round(mtd, 2),
            "kuverter": kuverter,
            "kuvert_budget": round(kuverter * kp, 2),
            "ekstrasalg_inkl": round(ekstra_inkl, 2),
            "vareforbrug": round(vareforbrug, 2),
            "loebende_budget": round(budget_cum + vare_cum, 2),
            "diff": round(budget_cum + vare_cum - mtd, 2),
            "kostpris_pr_kuvert": round((mtd - vare_cum) / kuv_cum, 2) if kuv_cum else None,
        })
    return dage
