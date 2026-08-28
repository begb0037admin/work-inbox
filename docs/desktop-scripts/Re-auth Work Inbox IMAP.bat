@echo off
REM ===================================================================
REM  Re-auth Work Inbox IMAP   (reference copy of the Desktop script)
REM  Live location: D:\OneDrive - lelitte.com\Desktop\Re-auth Work Inbox IMAP.bat
REM
REM  Run this:
REM    - ONCE, before MAIL_BACKEND=imap is ever used, to prime the token cache;
REM    - AGAIN whenever the "IMAP mail sign-in expired" toast fires.
REM
REM  It pulls imap_mail.py + reauth_imap.py fresh from main into the run dir
REM  (same mechanism the briefing .bat uses for fetch_inbox.py -- no stale
REM  Desktop/clone copy), then runs the device-code sign-in. You get a short
REM  code + a URL; approve in a browser on any device. Nothing is sent. No
REM  secret is stored (the client id is Mozilla Thunderbird's public one).
REM  Token cache -> %LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin
REM ===================================================================
setlocal
set "PROJECT_DIR=C:\Users\admin\Documents\Claude\Projects\work-inbox"
set "BASE=https://raw.githubusercontent.com/begb0037admin/work-inbox/main"
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMddHHmmss"') do set "CB=%%I"

if not exist "%PROJECT_DIR%\" (
  echo ERROR: run dir not found: %PROJECT_DIR%
  pause
  exit /b 1
)

echo.
echo [%DATE% %TIME%] Pulling imap_mail.py + reauth_imap.py from main...
curl -fsSL "%BASE%/imap_mail.py?cb=%CB%"   -o "%PROJECT_DIR%\imap_mail.py"   || goto :pullfail
curl -fsSL "%BASE%/reauth_imap.py?cb=%CB%" -o "%PROJECT_DIR%\reauth_imap.py" || goto :pullfail
echo [%DATE% %TIME%] OK.

cd /d "%PROJECT_DIR%"
echo.
python "%PROJECT_DIR%\reauth_imap.py"
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
  echo [%DATE% %TIME%] Re-auth complete. The scheduled briefing will refresh this token silently until it next expires.
) else (
  echo [%DATE% %TIME%] Re-auth did NOT complete cleanly ^(exit %RC%^). See the messages above.
)
echo.
pause
endlocal
exit /b %RC%

:pullfail
echo [%DATE% %TIME%] ERROR: could not download the scripts from GitHub. Check network / raw.githubusercontent.com.
pause
endlocal
exit /b 1
