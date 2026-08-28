@echo off
REM ===================================================================
REM  Run Mail Parity Test   (reference copy of the Desktop script)
REM  Live location: D:\OneDrive - lelitte.com\Desktop\Run Mail Parity Test.bat
REM
REM  Captures the Phase 1 MAIL pull BOTH ways (Outlook COM and IMAP+OAuth2)
REM  in the same window and diffs them field-by-field on the internet
REM  Message-ID. Phase-1-only: NO Granola, NO calendar, NO AI call, NO push,
REM  NO ledger / Command-Centre writes. Each capture is ~10-20s.
REM
REM  PRECONDITIONS:
REM    1. Run "Re-auth Work Inbox IMAP.bat" once first (primes the IMAP token).
REM    2. Classic Outlook running + "Connected to: Microsoft Exchange"
REM       (only the COM half's mail pull needs it -- the calendar dep is gone).
REM
REM  OUTPUT (under C:\Users\admin\Documents\Claude\Projects\work-inbox):
REM    data\parallel\{com,imap}_{inbox,sent}_raw.json
REM    data\parallel\parity_<timestamp>.json    (real_issues / readcap_churn / benign)
REM    mail_parity_last_run.log                 (full stdout+stderr of this whole run)
REM ===================================================================
setlocal
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PROJECT_DIR=C:\Users\admin\Documents\Claude\Projects\work-inbox"
set "BASE=https://raw.githubusercontent.com/begb0037admin/work-inbox/main"
set "LOG=%PROJECT_DIR%\mail_parity_last_run.log"
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMddHHmmss"') do set "CB=%%I"

if not exist "%PROJECT_DIR%\" (
  echo ERROR: run dir not found: %PROJECT_DIR%
  pause & exit /b 1
)
cd /d "%PROJECT_DIR%"

REM Everything (pulls + both captures + diff) is appended to %LOG%.
echo Run Mail Parity Test - %DATE% %TIME%> "%LOG%"

echo Pulling fetch_inbox.py + imap_mail.py + diff_mail_pull.py from main...
curl -fsSL "%BASE%/fetch_inbox.py?cb=%CB%"    -o "%PROJECT_DIR%\fetch_inbox.py"    >> "%LOG%" 2>&1 || goto :pullfail
curl -fsSL "%BASE%/imap_mail.py?cb=%CB%"      -o "%PROJECT_DIR%\imap_mail.py"      >> "%LOG%" 2>&1 || goto :pullfail
curl -fsSL "%BASE%/diff_mail_pull.py?cb=%CB%" -o "%PROJECT_DIR%\diff_mail_pull.py" >> "%LOG%" 2>&1 || goto :pullfail

REM AI backend matches the live pipeline, but the parity path exits right
REM after the mail dump (before any AI call), so it is never actually reached.
set "AI_BACKEND=claude_code"
set "ANTHROPIC_API_KEY="

echo ===== CAPTURE 1/2  MAIL_BACKEND=com =====>> "%LOG%"
echo CAPTURE 1/2  MAIL_BACKEND=com ...
set "MAIL_BACKEND=com"
set "WI_MAIL_PARALLEL=1"
python -u "%PROJECT_DIR%\fetch_inbox.py" >> "%LOG%" 2>&1
echo   com capture exit %ERRORLEVEL%

echo ===== CAPTURE 2/2  MAIL_BACKEND=imap =====>> "%LOG%"
echo CAPTURE 2/2  MAIL_BACKEND=imap ...
set "MAIL_BACKEND=imap"
set "WI_MAIL_PARALLEL=1"
python -u "%PROJECT_DIR%\fetch_inbox.py" >> "%LOG%" 2>&1
echo   imap capture exit %ERRORLEVEL%

echo ===== DIFF =====>> "%LOG%"
echo DIFF ...
set "MAIL_BACKEND="
set "WI_MAIL_PARALLEL="
python -u "%PROJECT_DIR%\diff_mail_pull.py" >> "%LOG%" 2>&1
set "DIFF_RC=%ERRORLEVEL%"

echo.
echo ---------------- last 45 lines of %LOG% ----------------
powershell -NoProfile -Command "Get-Content -LiteralPath '%LOG%' -Encoding UTF8 | Select-Object -Last 45"
echo -------------------------------------------------------
echo.
echo  diff exit %DIFF_RC%   (0 = mail pulls match; 1 = real differences - read above)
echo  full log : %LOG%
echo  raw dumps + parity JSON : %PROJECT_DIR%\data\parallel\
echo.
pause
endlocal
exit /b %DIFF_RC%

:pullfail
echo [%DATE% %TIME%] ERROR: could not download scripts from GitHub. See %LOG%.
pause
endlocal
exit /b 1
