@echo off
REM ===================================================================
REM  Re-auth Work Inbox IMAP
REM  Run this when the "Outlook mail sign-in expired" toast fires (or
REM  once, before MAIL_BACKEND=imap is first used).
REM  Opens a device-code sign-in: you get a short code + a URL, approve
REM  in a browser on any device, done. Nothing is sent. No secret stored.
REM  Safe to run any time; it only refreshes the read-only mail token.
REM ===================================================================
setlocal
cd /d "%~dp0"

echo.
echo  Pulling the latest reauth_imap.py / imap_mail.py from GitHub...
git fetch origin --quiet 2>nul
git checkout origin/main -- reauth_imap.py imap_mail.py 2>nul

echo.
python "%~dp0reauth_imap.py"
set RC=%ERRORLEVEL%

echo.
if "%RC%"=="0" (
  echo  Re-auth complete. You can close this window.
) else (
  echo  Re-auth did NOT complete cleanly ^(exit %RC%^). See the messages above.
)
echo.
pause
endlocal
