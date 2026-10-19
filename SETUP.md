# Technician Log: update steps

1. **Google script**: Apps Script -> replace all code with `Code.gs` (secret is already GST12345) -> Deploy -> Manage deployments -> pencil -> Version: New version -> Deploy.
2. **config.js**: paste your `SCRIPT_URL` (ends in /exec). `KEY` must equal `SECRET` in Code.gs. Staff details (Emp Code, designation) are already inside.
3. **Netlify**: Deploys tab -> drag the whole unzipped `technician-log` folder (not the zip). Check `your-site/sw.js` first line says `techlog-v8`.
   - **Backup (once):** in Apps Script choose the function `setupBackup` in the toolbar and press Run, then approve the Drive permission. A copy of your sheet is saved daily in the Drive folder "Technician Log Backups" (last 14 kept).
   - **PINs (once):** in Apps Script choose the function `setupPins` and press Run. It fills the Lists tab with the technicians, their PINs and the admin PIN. Later you can change any PIN directly in the Lists tab (column B for technicians, cell D2 for the admin). PINs are up to 6 digits; leave one empty to skip that person.
   - **(older note) PINs:** in the sheet's **Lists** tab, type each technician's PIN in column B (next to the name). Format column B as Plain text. Leave a PIN empty to skip it for that person. Once any PIN is set, phones ask each technician to sign in once.
   - **Utility meters:** the first time a phone loads, a **Meters** tab (your 7 meters) and later a **Readings** tab appear in the sheet. To add or change a meter, edit the Meters tab (Type, Unit/Area, Meter No, Account No). Format those columns as Plain text.
4. **Phones**: open the app twice while online to update.
5. **Admin page** (`your-site/admin.html`): sign in with the admin PIN, then pick dates ->
   - Download daily log (Excel): one file, one sheet per day, Day/Night shift chosen on the page
   - Download overtime (Excel): one file, one sheet per technician (per month), Compatibility Report last
   - Download utility readings (Excel): Summary, Electricity, Water and Daily Readings sheets with consumption (today minus previous reading)
   - Performance dashboard: jobs, completion %, job hours, overtime hours per technician
