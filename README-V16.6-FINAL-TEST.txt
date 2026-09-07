PANCHAYAT WORK GPS V16.6 — FINAL TEST BUILD

FIXES:
- XLSX import: converts upload blob to application/zip before Utilities.unzip, fixing the ContentType error.
- Admin Delete All clears all 15 work columns, so Total Works becomes 0 immediately after refresh.
- Excel upload repopulates WorkGPS and Dashboard count refreshes.
- Capture shows Photo 1 / 4, 2 / 4, 3 / 4, 4 / 4 progress.
- Quick Save success message explicitly shows the photo number.
- Fourth photo completes the work and hides it from normal User Capture list.
- Fifth photo is blocked by existing backend nextPhoto_ and frontend filtering.
- Existing V16.5 IndexedDB/background sync, Sansad filtering, Admin Settings, User Activity and theme system preserved.

DEPLOY:
GAS: Replace Code.gs, Save, Deploy > Manage deployments > Edit > New version > Deploy.
Cloudflare: Replace public/index.html in the Worker project. Keep current worker.js/wrangler.jsonc unless separately changed. Then run npx.cmd wrangler deploy.

TEST: Admin login -> Database -> Delete All -> Dashboard must show 0. Then Import XLSX -> Dashboard count returns. User -> Sansad -> Work -> Photo 1/2/3/4. After Photo 4 the work disappears from User Capture. Settings visible only for Admin.
