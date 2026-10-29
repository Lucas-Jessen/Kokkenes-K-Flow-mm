# Opsætning: mail → Excel med Power Automate

Hver gang der kommer en "Ordrekvittering" i `indkoebkk@forenede.dk`, læser flowet mailen og skriver ordren
ind i `K-flow.xlsx` på SharePoint. Månedsarket regner selv videre (formlerne henter fra tabellen *Ordrer*).

```
Ny mail i indkoebkk@  ──►  Power Automate  ──►  Office Script "Indlæs ordrekvittering"  ──►  K-flow.xlsx
                                                  (læser mailen, finder kolonne,           ├ Ordrer
                                                   springer dubletter over)                ├ Ordrelinjer
                                                                                           └ Oktober 2026 … (formler)
```

Kræver: adgang til den delte postkasse, Power Automate og Excel på nettet (standard i Microsoft 365 – ingen premium-licens).

## 1. Læg arbejdsbogen på SharePoint

1. Lav en tom arbejdsbog med de måneder, der skal bruges:
   ```bash
   python koer.py skabelon --fra 2026-11 --til 2027-12 --excel output/K-flow.xlsx
   ```
2. Upload `K-flow.xlsx` til et SharePoint-dokumentbibliotek, som køkkenet har adgang til (fx *Køkken › K-flow*).
3. Tjek fanen **Kolonneregler**: den bestemmer, hvilken kolonne (konto) en ordre tæller i. Den kan rettes direkte i Excel.

## 2. Gem Office Scriptet

1. Åbn `K-flow.xlsx` i **Excel på nettet** → fanen **Automatiser** → **Nyt script**.
2. Slet indholdet, og indsæt hele `office-scripts/IndlaesOrdrekvittering.ts`.
3. Omdøb scriptet til **Indlæs ordrekvittering**, og gem.

## 3. Byg flowet

[make.powerautomate.com](https://make.powerautomate.com) → **Opret** → **Automatiseret cloudflow**.

| # | Handling | Indstillinger |
|---|---|---|
| 1 | **Når der kommer en ny mail i en delt postkasse (V2)** (Office 365 Outlook) | Postkasse: `indkoebkk@forenede.dk` · Mappe: Indbakke · *Emnefilter*: `Ordrekvittering` · Kun med vedhæftede filer: Nej |
| 2 | **Kør script** (Excel Online (Business)) | Placering: SharePoint-webstedet · Dokumentbibliotek · Fil: `K-flow.xlsx` · Script: *Indlæs ordrekvittering* · **mailHtml** = *Brødtekst* (Body) · **emne** = *Emne* · **modtaget** = *Modtaget tidspunkt* |
| 3 | **Betingelse** | `body/result/status` *er ikke lig med* `indlæst` **og** *er ikke lig med* `dublet` |
| 3a | Hvis ja → **Send en mail (V2)** | Til: dig selv/køkkenchefen · Emne: `K-flow: mail kunne ikke indlæses` · Brødtekst: status, ordre, advarsler fra trin 2 + link til mailen |
| 4 | (valgfrit) **Flyt mail** | Flyt behandlede mails til mappen *Indlæst*, så indbakken er overskuelig |

Indstillinger på trigger-handlingen → **Samtidighedskontrol: Til, grad 1**, så to mails ikke skriver i arket på samme tid.

> Ældre mails, der allerede ligger i postkassen, tages ikke med af triggeren. De kan gemmes som `.eml` i en mappe og indlæses med
> `python koer.py importer-mails <mappe> --excel <sti til K-flow.xlsx>` (mens filen ikke er åben hos andre).

## 4. Test

1. Videresend den rigtige ordrekvittering (ordre 31525653) til `indkoebkk@forenede.dk` med emnet uændret.
2. Følg kørslen under **Mine flows › Kørselsoversigt**. Resultatet af *Kør script* skal vise `"status": "indlæst"`.
3. I `K-flow.xlsx`: én ny række i **Ordrer**, 5 i **Ordrelinjer**, og 386,23 i **Frugt** den 27. i månedsarket.
4. Send samme mail igen → status `dublet`, og intet nyt i arket.

Hvis status er `fejl` med "Ingen varelinjer fundet", er den rigtige mails HTML bygget anderledes end vores eksempel.
Send mig mailen som `.eml`, så tilpasser jeg scriptet (og tilføjer den som test i `tests/`).

## 5. Dashboardet

`python koer.py dashboard --excel "<OneDrive-synkroniseret sti til K-flow.xlsx>"` læser arket og skriver `web/data.js`.
I piloten kan det køres af Windows Opgavestyring hver time. Senere lægges hjemmesiden et sted med Microsoft-login
(SharePoint-side eller Azure Static Web Apps), og opdateringen flyttes med derhen.
