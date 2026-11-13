# Catalogul Steam în GitHub Actions

1. Creează un repository GitHub privat și copiază `steam_catalog.py` la rădăcină.
2. Creează fișierul `.github/workflows/steam-catalog.yml` pe ramura implicită și
   copiază în el conținutul fișierului cu același nume din această arhivă.
3. În **Settings → Secrets and variables → Actions → New repository secret**,
   creează `STEAM_API_KEY` și lipește cheia în câmpul **Secret**.
4. În **Actions → Catalog Steam**, apasă **Run workflow** o singură dată.
   Poți închide browserul și calculatorul după ce ai pornit rularea.
5. Deschide ulterior rularea și descarcă arhiva **steam-catalog** din
   **Artifacts**. Ea conține `steam_games.xlsx` și `steam_progress.sqlite3`.

Exportul folosește modul rapid, include jocurile cu data lansării de la 1
ianuarie 2024 și afișează numărul de recenzii folosite în scorul magazinului.
Scriptul colectează date în loturi și poate lăsa câmpul de recenzii gol dacă
Steam nu returnează numărul. Prima rulare pornește cu un catalog nou.

Dacă operațiunea ajunge la limita internă de 5 ore și 25 de minute, scriptul
primește semnalul de oprire și exportă datele parțiale înainte de limita jobului.
Pentru reluare, deschide o nouă rulare și completează câmpul **resume_run_id**
cu numărul din adresa rulării precedente, după `/actions/runs/`. Acea rulare
trebuie să aibă încă arhiva **steam-catalog**, disponibilă 14 zile. Un job
întrerupt abrupt înainte de încărcarea arhivei poate pierde progresul acelei rulări.

Repository-urile private consumă minute GitHub Actions din cota contului.

## Steam Next Fest · octombrie 2026

După ce workflow-ul ajunge pe ramura implicită, deschide **Actions → Steam Next
Fest — October 2026 → Run workflow**. Descarcă arhiva
**steam-next-fest-october-2026** din **Artifacts**. Arhiva conține raportul
`steam_next_fest.xlsx` și baza de progres pentru reluare prin `resume_run_id`.
Colectarea nu necesită o cheie Steam API.

Raportul include câte un rând pentru fiecare dintre cele 7.319 AppID-uri
din [pagina oficială a evenimentului Steam](https://store.steampowered.com/sale/nextfest_oct2026),
inclusiv jocurile care încă nu s-au lansat. Datele reflectă momentul verificării,
indicat în UTC pentru fiecare sursă. `review_count` este numărul de recenzii
folosite în scorul afișat de magazin pentru jocul de bază (`store_filtered`).
Un zero este valoarea raportată de Steam. Un câmp gol înseamnă că numărul nu
a putut fi verificat (`unavailable`) sau că verificarea este încă în așteptare
(`pending`).

`ai_disclosure_status` este `present` când Steam afișează declarația,
`absent` când pagina jocului este accesibilă și secțiunea lipsește,
`unavailable` când pagina sau textul declarației nu a putut fi citit și
`pending` când verificarea este încă în așteptare. Textul declarației este în
`ai_disclosure`. Paginile se verifică în engleză, cu România ca regiune implicită.

Pentru a rula local, folosește Python 3.10 sau mai nou:

```bash
python -m pip install openpyxl==3.1.5
python steam_next_fest.py
```

Repetă comanda cu aceeași bază de progres pentru a verifica rândurile rămase și
a reîncerca datele indisponibile. Pentru o colectare nouă, specifică un alt fișier
`--db`. Poți folosi `--country` pentru altă regiune și `--ids` pentru alt CSV
cu coloana `app_id`.
