"""Læser mails i den fælles postkasse via Microsoft Graph i stedet for Outlook – ingen pop-ups, Outlook skal ikke køre.

Kræver en app-registrering i Entra ID med adgang til postkassen (se README, afsnittet "Microsoft Graph").
Oplysningerne ligger i graph.json ved siden af dette script – filen er i .gitignore og må aldrig committes:

  {"tenant_id": "...", "client_id": "...", "client_secret": "..."}

Findes graph.json ikke, bruger ordre_til_ark.py Outlook som før.

  python graph_mail.py        # test: hent token og vis de seneste mails fra AB (markerer intet som læst)
"""
import json
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

INDSTILLINGER = Path(__file__).resolve().parent / "graph.json"
GRAPH = "https://graph.microsoft.com/v1.0"


def konfigureret():
    return INDSTILLINGER.exists()


def _indstillinger():
    d = json.loads(INDSTILLINGER.read_text(encoding="utf-8"))
    mangler = [k for k in ("tenant_id", "client_id", "client_secret") if not d.get(k)]
    if mangler:
        raise RuntimeError(f"graph.json mangler: {', '.join(mangler)}")
    return d


def token():
    d = _indstillinger()
    krop = urllib.parse.urlencode({"grant_type": "client_credentials", "client_id": d["client_id"],
                                   "client_secret": d["client_secret"],
                                   "scope": "https://graph.microsoft.com/.default"}).encode()
    url = f"https://login.microsoftonline.com/{d['tenant_id']}/oauth2/v2.0/token"
    try:
        with urllib.request.urlopen(urllib.request.Request(url, data=krop), timeout=60) as r:
            return json.load(r)["access_token"]
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Kunne ikke logge på Microsoft Graph: {e.read().decode(errors='replace')[:300]}") from None


def _kald(metode, url, tok, data=None):
    req = urllib.request.Request(url, method=metode, data=json.dumps(data).encode() if data is not None else None,
                                 headers={"Authorization": f"Bearer {tok}", "Content-Type": "application/json",
                                          "Prefer": 'outlook.body-content-type="html"'})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            svar = r.read()
            return json.loads(svar) if svar else None
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Graph svarede {e.code} på {metode} {url.split('?')[0]}: "
                           f"{e.read().decode(errors='replace')[:300]}") from None


def mails(postkasse, afsender, dage_tilbage):
    """Giver (html, modtaget, emne, markér_læst) for mails fra – eller videresendt fra – `afsender`."""
    tok = token()
    fra = (datetime.now(timezone.utc) - timedelta(days=dage_tilbage)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url = (f"{GRAPH}/users/{urllib.parse.quote(postkasse)}/mailFolders/inbox/messages?"
           + urllib.parse.urlencode({"$filter": f"receivedDateTime ge {fra}", "$orderby": "receivedDateTime desc",
                                     "$select": "id,subject,receivedDateTime,from,body,isRead", "$top": "50"}))
    while url:
        side = _kald("GET", url, tok)
        for m in side.get("value", []):
            fra_adr = ((m.get("from") or {}).get("emailAddress") or {}).get("address", "")
            html = (m.get("body") or {}).get("content", "")
            if afsender.lower() not in fra_adr.lower() and afsender.lower() not in html.lower():
                continue
            modtaget = datetime.fromisoformat(m["receivedDateTime"].replace("Z", "+00:00")).astimezone()
            yield html, modtaget.strftime("%Y-%m-%d %H:%M"), m.get("subject", ""), _markoer(postkasse, m, tok)
        url = side.get("@odata.nextLink")


def _markoer(postkasse, m, tok):
    def markér_læst():
        if not m.get("isRead"):
            _kald("PATCH", f"{GRAPH}/users/{urllib.parse.quote(postkasse)}/messages/{m['id']}", tok, {"isRead": True})
    return markér_læst


if __name__ == "__main__":
    from ordre_til_ark import AFSENDER, DAGE_TILBAGE, POSTKASSE
    token()
    print("Login til Microsoft Graph virker.")
    for i, (_, modtaget, emne, _) in enumerate(mails(POSTKASSE, AFSENDER, DAGE_TILBAGE)):
        if i == 10:
            break
        print(f"  {modtaget}  {emne}")
