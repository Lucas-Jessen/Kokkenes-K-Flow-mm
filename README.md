# Kokkenes K-flow

Henter ordrekvitteringer fra AB Catering (kundeservice@ab-kbh.dk) i den fælles postkasse
indkoebkk@forenede.dk via klassisk Outlook og skriver dem i `Kokkenes k ark.xlsx`.

- **Ordreliste**: én række pr. varelinje. Ordrenummeret bruges som dubletkontrol.
- **Januar–December**: kolonnen *BC catering* summerer Ordreliste pr. leveringsdag med SUMIFS.
- Behandlede mails markeres som læst.

## Kørsel

```
python ordre_til_ark.py              # hent nye mails fra Outlook
python ordre_til_ark.py --opsaet     # opret Ordreliste og BC-formler i arket
python ordre_til_ark.py --mappe DIR  # læs .msg/.eml/.html fra en mappe i stedet
python ordre_til_ark.py --vis-postkasser
```

`Hent ordrer.bat` kører scriptet manuelt og logger til `ordre_til_ark.log`.
En planlagt opgave i Windows ("K-flow - Hent ordrer fra AB Catering") kører det hver time med `pythonw.exe`.

Kræver Windows, klassisk Outlook, Excel og `pip install pywin32`.

Excel-filer og log ligger kun lokalt og er ikke med i repoet.
