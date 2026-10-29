# Køkkenets K-flow

Pilot: ordrekvitteringer fra **AB Catering København** læses automatisk fra indbakken
`indkoebkk@forenede.dk`, lægges i Excel (samme opsætning som det manuelle månedsark) og vises på en
hjemmeside sammen med Kanpla-tal (kuverter + ekstrasalg), så man kan se, om køkkenet er i plus eller minus mod budget.

```
indkoebkk@forenede.dk ──► (Power Automate) gem mail i mappe ──► koer.py ──► output/K-flow.xlsx
                                                                   ▲   └──► web/data.js ──► web/index.html (dashboard)
                                    Kanpla-eksport (CSV) ──────────┘
```

## Kom i gang

```bash
pip install openpyxl
python scripts/lav_demodata.py   # laver demo-mails + Kanpla-fil for oktober 2026
python koer.py                   # læser mails og Kanpla, bygger Excel + dashboard-data
python -m http.server 8765 --directory web   # åbn http://localhost:8765
python -m unittest -v            # tester mail-læseren
```

## Hvad ligger hvor

| Fil | Indhold |
|---|---|
| `config.json` | Kantine, kuvertpris, kolonner (konti), leverandører og regler for hvilken kolonne en ordre lander i, ekstrasalgs-satser |
| `kkflow/parse_ordre.py` | Læser "Ordrekvittering"-mails (.eml / .html / .txt): ordre, leveringsdato, kunde, bruger, leveringsbemærkning, varelinjer, total. Advarer hvis sum af linjer ≠ total |
| `kkflow/kanpla.py` | Indlæser Kanpla-eksport (kuverter + ekstrasalg pr. dag) |
| `kkflow/beregning.py` | Samme regnestykke som det manuelle ark (bruges til dashboardet) |
| `kkflow/excel_export.py` | Bygger Excel-filen |
| `koer.py` | Kører det hele. Ordrer gemmes i `data/ordrer.json` pr. ordrenummer, så samme mail aldrig tælles to gange |
| `web/index.html` | Dashboardet (ingen eksterne afhængigheder) |
| `eksempler/` | Demo-data. Ordre 31525653 er gengivet fra den rigtige mail – resten er opdigtet |

## Excel-filen

* **Oktober 2026** (ét ark pr. måned) – opbygget som jeres ark: dato-rækker, en kolonne pr. konto/kategori
  (BC catering, Slagter & Fisk, Kaffe/The, Mælk, Frugt, Arrangementer, Carlsberg), Indkøb i alt, MTD,
  Øvrige omkostninger, Kuverter, Kuvert budget, Kuvertpris, Løbende budget, Diff fra budget og ekstrasalg med vareforbrugssatser.
  Blå celler hentes med formler fra arkene nedenfor; gule celler indtastes stadig manuelt.
* **Ordrer** – én række pr. ordre.
* **Ordrelinjer** – én række pr. varelinje (med ordrens hovedfelter på hver linje), klar til filtrering/pivot.
* **Kanpla** – kuverter og ekstrasalg pr. dag.

Regnestykket (genskabt fra jeres ark og kontrolleret mod tallene i udklippet):

* Kuvert budget = kuverter × kuvertpris (kr. 25,94)
* Vareforbrug ekstrasalg = Σ(beløb inkl. moms × vareforbrugssats) / 1,25 (Fredagsmorgenmad = antal × 25 kr.)
* Løbende budget = Σ kuvert budget + Σ vareforbrug
* Diff fra budget = løbende budget − MTD-indkøb (**plus = under budget**)
* Kuvertpris (faktisk) = (MTD − Σ vareforbrug) / Σ kuverter – grøn hvis ≤ 25,94, ellers rød

## Mod drift (forslag)

1. **Mail ind:** Power Automate-flow på den delte postkasse `indkoebkk@forenede.dk`: "Når en ny mail ankommer"
   med emne *Ordrekvittering* → gem mailen som `.eml` i en SharePoint/OneDrive-mappe (`indbakke`).
2. **Kørsel:** `koer.py --indbakke <mappe> --kanpla <mappe>` køres automatisk (fx hver time) – på en server,
   som Azure Function, eller vi flytter logikken direkte ind i Power Automate/Office Scripts.
3. **Kanpla:** daglig eksport (CSV) eller API, hvis Kanpla giver adgang. Kolonnenavne tilpasses i `kkflow/kanpla.py`.
4. **Hjemmeside til hele firmaet:** `web/` er statiske filer. Kan lægges på SharePoint/Teams-fane eller
   Azure Static Web Apps med Microsoft-login, så kun medarbejdere i Forenede har adgang.

## Åbne spørgsmål

* Hvilken kolonne skal AB Catering-ordrer stå i? Nu: kunde "… FRUGT" → **Frugt (21608)**, alt andet → **BC catering (21602)**.
* Er beløbene i mailen ex moms? (antaget ja, som "Ex moms" i arket)
* Tæller "Øvrige omkostninger" med i "Indkøb i alt"? (antaget ja)
* Format på Kanpla-eksporten (et rigtigt eksempel ville hjælpe) og en rigtig `.eml` fra indbakken.
