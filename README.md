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
