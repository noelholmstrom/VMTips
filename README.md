# VM-tips 2026 — automatisk ställning

En liten webbsida som visar ställningen i tipstävlingen och som **uppdaterar
matchresultaten automatiskt**. Poängen räknas fortfarande ut i ditt vanliga
Excel-ark (gruppspel, slutspelsbonus och bonusfrågor) — automationen fyller bara
i resultaten, räknar om arket och publicerar ställningen.

## Så funkar det

```
football-data.org  ──►  update_results.py  ──►  räknar om arket (LibreOffice)
                                                        │
                                                        ▼
                              data.json  ◄── exporteras ──┘
                                  │
                                  ▼
   index.html (GitHub Pages)  läser data.json och visar tabellen
```

GitHub kör `update_results.py` enligt schema (var 30:e min). Inget behöver göras
manuellt under turneringen.

## Filer

| Fil | Vad |
|-----|-----|
| `index.html` | Själva sidan (tabell + matcher + varje spelares tips). |
| `data.json` | Genereras automatiskt — ställning och alla tips. |
| `VM-tips-2026-Swedish_7p_4_4_MASTER.xlsm` | Ditt tipsark — sanningskällan för poängen. |
| `update_results.py` | Hämtar resultat, skriver i arket, räknar om, exporterar `data.json`. |
| `recalc.py` | Räknar om arket med LibreOffice. |
| `manual_results.json` | (valfritt) Handpåläggning om API:t skulle ha fel/sakna en match. |
| `.github/workflows/update.yml` | Schemat som kör allt. |

## Uppsättning (engångs, ~10 min)

1. **Skapa ett repo på GitHub** och lägg upp alla filerna ovan
   (`update.yml` ska ligga i `.github/workflows/`).

2. **Hämta en gratis API-nyckel** på <https://www.football-data.org/client/register>.
   VM (competition `WC`) ingår i gratisnivån.

3. **Lägg in nyckeln som secret:** repo → *Settings* → *Secrets and variables* →
   *Actions* → *New repository secret*.
   Namn: `FOOTBALL_DATA_TOKEN`, värde: din nyckel.

4. **Slå på GitHub Pages:** *Settings* → *Pages* → *Source: Deploy from a branch* →
   `main` / `/ (root)*`. Sidan ligger sen på `https://<användarnamn>.github.io/<repo>/`.

5. **Testkör flödet:** *Actions* → *Uppdatera VM-resultat* → *Run workflow*.
   När det är klart har `data.json` uppdaterats och sidan visar senaste ställningen.

Det var allt — sen sköter schemat resten.

## Om något skulle strula

* **En match får fel/inget resultat från API:t.** Lägg in den för hand i
  `manual_results.json` med matchnumret som nyckel (samma nummer som i arket):

  ```json
  { "7": [1, 1], "25": [1, 0] }
  ```

  Manuella resultat vinner alltid över API:t. Commita filen så fyller nästa
  körning i dem.

* **Nya tippare eller ändrade tips.** Uppdatera arket i Excel som vanligt och
  commita den nya `.xlsm`-filen. Nästa körning plockar upp ändringarna.
  (Tipparnas namn i `update_results.py` måste matcha flik-namnen i arket.)

* **Köra lokalt för test:** `python update_results.py --mock mock.json`
  använder en lokal JSON istället för att anropa API:t.

## Poängsystem (oförändrat från arket)

* Gruppmatch: 2 p per lag med rätt antal mål + 3 p för rätt tecken = max **7 p**.
* Slutspel: 1 p (32-del) upp till 8 p (final) per rätt lag vidare.
* Bonusfrågor: 20 p styck (Skyttekung, Världsmästare m.fl.).
