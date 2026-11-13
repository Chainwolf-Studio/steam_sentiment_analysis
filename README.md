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

## Recenzii pentru jocurile Next Fest filtrate

Lista implicită conține cele 58 de AppID-uri distincte din tabelul
`ClasificareLocalizare` al fișierului `steam_next_fest_clasificat.xlsx`.
Workflow-ul **Steam Next Fest — recenzii jocuri filtrate** se poate porni din
**Actions → Run workflow** după ce ajunge pe ramura implicită. Rularea nu
necesită o cheie API Steam.

- `reviews.csv` are un rând pentru fiecare recenzie publică a jocului de bază,
  cu AppID, numele și genurile Steam, ID-ul recenziei, valența, codul de limbă,
  data creării ca timestamp Unix și textul. Valența este recomandarea Steam
  (`positive` sau `negative`). Genurile multiple sunt delimitate cu ` | `.
- `games.csv` arată progresul pentru fiecare joc, inclusiv jocurile fără
  recenzii. `expected_reviews` este valoarea `total_matching` raportată de
  Steam la începutul paginării. O diferență este marcată ca `count_mismatch`.
- `reviews.sqlite3` păstrează textul original și progresul după fiecare pagină.
  În CSV, valorile ce ar putea fi interpretate ca formule primesc un apostrof
  la început.

Colectarea include toate limbile, ambele valențe, toate tipurile de achiziție
și recenziile din perioade marcate off-topic. Nu există un filtru de dată
pentru Next Fest. Genurile provin din metadatele magazinului în engleză, cu
România ca regiune implicită.

Pentru rulare locală:

```bash
python -m pip install openpyxl==3.1.5
python steam_reviews.py
```

Rulează din nou aceeași comandă pentru a relua progresul sau folosește
`resume_run_id` în GitHub Actions. Pentru o colectare nouă folosește alt
`--output-dir`. Poți schimba lista cu `--ids` și un CSV cu coloana `app_id`.
