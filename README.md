# Køkkenets K-flow

Omhandler projektet der laves omkring opsætning af indhentning af data fra mail, til excel ark.

Pilot: ordrekvitteringer fra **AB Catering København**, der kommer ind i `indkoebkk@forenede.dk`, skrives automatisk
ind i Excel-arket `K-flow.xlsx` (samme opsætning som det manuelle månedsark). Arket vises på en hjemmeside sammen med
Kanpla-tal (kuverter og ekstrasalg), så man kan se, om køkkenet er i plus eller minus mod budget.

```
indkoebkk@forenede.dk ─► Power Automate ─► Office Script ─► K-flow.xlsx (SharePoint) ─► koer.py dashboard ─► web/
                                                                 ▲
                                          Kanpla-eksport ────────┘ (koer.py importer-kanpla)
```

**Opsætning af flowet: se [docs/power-automate-opsaetning.md](docs/power-automate-opsaetning.md).**

## Kom i gang (demo)

```bash
pip install openpyxl
python scripts/lav_demodata.py   # demo-mails + Kanpla-fil for oktober 2026
python koer.py demo              # skabelon -> indlæs mails -> indlæs Kanpla -> dashboard (output/K-flow-demo.xlsx)
python -m http.server 8765 --directory web   # åbn http://localhost:8765
python -m unittest -q            # tester mail-læseren (Python)
node --test tests/office_script.test.mjs     # tester Office Scriptet
```

## Hvad ligger hvor

| Fil | Indhold |
|---|---|
| `office-scripts/IndlaesOrdrekvittering.ts` | **Det, der kører i drift.** Læser mailen og skriver ordren i tabellerne *Ordrer* og *Ordrelinjer*, springer dubletter over |
| `config.json` | Kantine, kuvertpris, kolonner (konti), start-regler for kolonner, ekstrasalgs-satser – bruges når skabelonen laves |
| `kkflow/parse_ordre.py` | Samme mail-læser i Python (til gamle mails/test) |
| `kkflow/arbejdsbog.py` | Læser/skriver tabellerne i K-flow.xlsx |
| `kkflow/kanpla.py` | Indlæser Kanpla-eksport (kuverter + ekstrasalg pr. dag) |
| `kkflow/beregning.py` | Samme regnestykke som månedsarket (til dashboardet) |
| `kkflow/excel_export.py` | Laver den tomme skabelon |
| `koer.py` | Kommandoer: `skabelon`, `importer-mails`, `importer-kanpla`, `dashboard`, `demo` |
| `web/index.html` | Dashboardet |
| `eksempler/` | Demo-data. Ordre 31525653 er gengivet fra den rigtige mail – resten er opdigtet |

## Excel-filen

* **Oktober 2026** (ét ark pr. måned) – opbygget som jeres ark: dato-rækker, en kolonne pr. konto/kategori
  (BC catering, Slagter & Fisk, Kaffe/The, Mælk, Frugt, Arrangementer, Carlsberg), Indkøb i alt, MTD,
  Øvrige omkostninger, Kuverter, Kuvert budget, Kuvertpris, Løbende budget, Diff fra budget og ekstrasalg med vareforbrugssatser.
  Blå celler hentes med formler fra arkene nedenfor; gule celler indtastes stadig manuelt.
* **Ordrer** – én række pr. ordre (skrives af Office Scriptet).
* **Ordrelinjer** – én række pr. varelinje (med ordrens hovedfelter på hver linje), klar til filtrering/pivot.
* **Kanpla** – kuverter og ekstrasalg pr. dag.
* **Kolonneregler** – hvilken leverandør/kunde der tæller i hvilken kolonne. Kan rettes direkte i Excel.

Regnestykket (genskabt fra jeres ark og kontrolleret mod tallene i udklippet):

* Kuvert budget = kuverter × kuvertpris (kr. 25,94)
* Vareforbrug ekstrasalg = Σ(beløb inkl. moms × vareforbrugssats) / 1,25 (Fredagsmorgenmad = antal × 25 kr.)
* Løbende budget = Σ kuvert budget + Σ vareforbrug
* Diff fra budget = løbende budget − MTD-indkøb (**plus = under budget**)
* Kuvertpris (faktisk) = (MTD − Σ vareforbrug) / Σ kuverter – grøn hvis ≤ 25,94, ellers rød

## Åbne spørgsmål

* Hvilken kolonne skal AB Catering-ordrer stå i? Nu: kunde "… FRUGT" → **Frugt (21608)**, alt andet → **BC catering (21602)**.
* Er beløbene i mailen ex moms? (antaget ja, som "Ex moms" i arket)
* Tæller "Øvrige omkostninger" med i "Indkøb i alt"? (antaget ja)
* Format på Kanpla-eksporten (et rigtigt eksempel ville hjælpe) og en rigtig `.eml` fra indbakken.
* Nye måneder: skabelonen laves for en periode ad gangen – skal scriptet selv oprette nye månedsark?
