# Kokkenes K-flow

Henter ordrekvitteringer fra AB Catering (kundeservice@ab-kbh.dk) i den fælles postkasse
indkoebkk@forenede.dk og skriver dem i `Kokkenes k ark.xlsx`. Spisende gæster hentes fra Kanpla,
og `rapport.html` viser indkøb mod budget.

- **Ordreliste**: én række pr. varelinje. Ordrenummeret bruges som dubletkontrol.
- **Kanpla**: spisende gæster pr. dag (HQ Ballerup). Månedsfanernes *Kuverter* henter derfra; *Kuvert manuel* overskriver.
- **Januar–December**: kolonnen *BC catering* summerer Ordreliste pr. leveringsdag med SUMIFS.
- Behandlede mails markeres som læst.

## Kørsel

```
python ordre_til_ark.py              # hent nye mails, Kanpla-gæster og lav rapport.html
python ordre_til_ark.py --opsaet     # opret Ordreliste og BC-formler i arket
python ordre_til_ark.py --mappe DIR  # læs .msg/.eml/.html fra en mappe i stedet
python kanpla.py --opdater           # hent Kanpla-gæster nu
python rapport.py                    # lav rapport.html nu
```

`Hent ordrer.bat` kører scriptet manuelt og logger til `ordre_til_ark.log`.
En planlagt opgave i Windows ("K-flow - Hent ordrer fra AB Catering") kører det ved logon og hver time med `pythonw.exe`.

Kræver Windows, Excel og `pip install pywin32`. Postkassen skal være tilføjet i klassisk Outlook; den læses med
Extended MAPI, så Outlook ikke viser sikkerheds-pop-ups og ikke behøver være åbent.

Excel-filer, log, rapport og nøgler ligger kun lokalt og er ikke med i repoet.
