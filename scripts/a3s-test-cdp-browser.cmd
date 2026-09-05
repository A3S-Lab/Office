@echo off
setlocal
bun "%~dp0a3s-test-cdp-browser.ts" %*
exit /b %ERRORLEVEL%
