Panchayat Work GPS V18.1 - Cloudflare Ready

1. Keep this entire folder together.
2. Google Apps Script backend: Google-Apps-Script-Backend/Code.gs
3. Deploy Code.gs separately in your existing Google Apps Script project.
4. Then double-click deploy-windows.bat from this folder.
5. Cloudflare serves ONLY files from public/; backend source is not exposed as a static asset.
6. Frontend calls /api and worker.js proxies that to the GAS endpoint configured in worker.js.
7. After deployment, hard refresh with Ctrl+Shift+R.
