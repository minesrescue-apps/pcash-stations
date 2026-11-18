# PCash Vs IP – Stations

MRS petty cash for Head Office and the stations (Carletonville, Evander, Welkom, Steelpoort) in one phone app.

- **Front end:** this repo (GitHub Pages). `config.json` holds the Apps Script `/exec` URL.
- **Back end:** `Code.gs` runs as a Google Apps Script web app under the MRS apps Google account.
  Data lives in the "PCash Vs IP – Stations" Google Sheet; slip photos in Drive under
  *Petty Cash Stations / <Station> / <yyyy-mm>*.
- **People:** the *Users* sheet (Name, PIN, Station code or ALL, Role MANAGER or ADMIN, Active Y/N).
  Managers see only their own station. Admins (finance) switch between stations.
- **Categories:** the *Stations* sheet. Head Office keeps PC / CF; stations have their own list.

No data, PINs or passwords are ever stored in this repo.
