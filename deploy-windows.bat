@echo off
title Panchayat Work GPS - Cloudflare Deploy
echo.
echo ==========================================
echo   Panchayat Work GPS - Cloudflare Deploy
echo ==========================================
echo.
echo Checking Node.js...
node --version
if errorlevel 1 (
  echo.
  echo ERROR: Node.js is not installed.
  echo Install Node.js LTS from https://nodejs.org/
  echo Then run this file again.
  pause
  exit /b 1
)
echo.
echo Checking Wrangler...
call npx wrangler --version
if errorlevel 1 (
  echo.
  echo ERROR: Wrangler could not start.
  pause
  exit /b 1
)
echo.
echo Opening Cloudflare login...
call npx wrangler login
if errorlevel 1 (
  echo.
  echo ERROR: Cloudflare login failed.
  pause
  exit /b 1
)
echo.
echo Deploying Panchayat Work GPS V18.2 ULTRA FINAL...
call npx wrangler deploy
if errorlevel 1 (
  echo.
  echo ERROR: Deployment failed.
  pause
  exit /b 1
)
echo.
echo ==========================================
echo   DEPLOYMENT COMPLETED
echo ==========================================
echo.
pause
