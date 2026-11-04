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
python graph_mail.py                 # test login til Microsoft Graph
python kanpla.py --opdater           # hent Kanpla-gæster nu
python rapport.py                    # lav rapport.html nu
```

`Hent ordrer.bat` kører scriptet manuelt og logger til `ordre_til_ark.log`.
En planlagt opgave i Windows ("K-flow - Hent ordrer fra AB Catering") kører det ved logon og hver time med `pythonw.exe`.

Kræver Windows, Excel og `pip install pywin32`. Uden Microsoft Graph kræves også klassisk Outlook.

Excel-filer, log, rapport og nøgler ligger kun lokalt og er ikke med i repoet.

## Microsoft Graph (mails uden Outlook og uden pop-ups)

Når `graph.json` findes, læses postkassen via Microsoft Graph i stedet for Outlook. Adgangen gives
med *RBAC for Applications* i Exchange Online, så appen **kun** kan læse indkoebkk@forenede.dk.

1. **Entra admin center** → *Applications* → *App registrations* → *New registration*
   - Name: `Kokkenes K-flow mail`, *Single tenant*, ingen redirect URI → *Register*.
   - Notér *Application (client) ID* og *Directory (tenant) ID*.
   - *Certificates & secrets* → *New client secret* → kopiér *Value*.
   - Giv **ikke** Mail-tilladelser under *API permissions* – de ville gælde alle postkasser.
2. **Enterprise applications** → find appen → notér dens *Object ID* (service principal).
3. **PowerShell som Exchange-administrator** (`Install-Module ExchangeOnlineManagement` første gang):

   ```powershell
   Connect-ExchangeOnline
   New-ServicePrincipal -AppId <client-id> -ObjectId <enterprise-app-object-id> -DisplayName "Kokkenes K-flow mail"
   New-ManagementScope -Name "Kun indkoebkk" -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'indkoebkk@forenede.dk'"
   New-ManagementRoleAssignment -App <client-id> -Role "Application Mail.ReadWrite" -CustomResourceScope "Kun indkoebkk"
   Test-ServicePrincipalAuthorization -Identity <client-id> -Resource indkoebkk@forenede.dk
   ```

   Den sidste linje skal vise `InScope : True`. Det kan tage op til en time, før adgangen virker.
4. Opret `graph.json` i projektmappen:

   ```json
   {"tenant_id": "...", "client_id": "...", "client_secret": "..."}
   ```

5. Kør `python graph_mail.py` – den viser de seneste AB-mails uden at markere noget som læst.

Sæt en påmindelse, før client secret udløber.
