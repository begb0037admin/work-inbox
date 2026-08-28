@echo off
REM ===================================================================
REM  Run Mail Parity Test   (reference copy of the Desktop script)
REM  Live location: D:\OneDrive - lelitte.com\Desktop\Run Mail Parity Test.bat
REM
REM  Captures the Phase 1 MAIL pull BOTH ways in the same window and diffs
REM  them field-by-field. Pushes NOTHING, mutates NOTHING (WI_MAIL_PARALLEL=1
REM  folds into the WI_AI_PARALLEL no-write posture). Does NOT flip any
REM  default -- MAIL_BACKEND is set only for this process, per capture.
REM
REM  PRECONDITIONS:
REM    1. Run "Re-auth Work Inbox IMAP.bat" once first (primes the IMAP token).
REM    2. Classic Outlook must be running + "Connected to: Microsoft Exchange"
REM       (the COM half needs it, exactly like the live briefing).
REM
REM  OUTPUT lands in:
REM    C:\Users\admin\Documents\Claude\Projects\work-inbox\data\parallel\
REM      com_inbox_raw.json  com_sent_raw.json
REM      imap_inbox_raw.json imap_sent_raw.json
REM      parity_<timestamp>.json   + a console report (also echoed below)
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
echo [%DATE% %TIME%] Pulling fetch_inbox.py + imap_mail.py + diff_mail_pull.py from main...
curl -fsSL "%BASE%/fetch_inbox.py?cb=%CB%"     -o "%PROJECT_DIR%\fetch_inbox.py"     || goto :pullfail
curl -fsSL "%BASE%/imap_mail.py?cb=%CB%"       -o "%PROJECT_DIR%\imap_mail.py"       || goto :pullfail
curl -fsSL "%BASE%/diff_mail_pull.py?cb=%CB%"  -o "%PROJECT_DIR%\diff_mail_pull.py"  || goto :pullfail
echo [%DATE% %TIME%] OK.

cd /d "%PROJECT_DIR%"

REM AI backend: match the live pipeline (subscription, no metered spend). The
REM mail dumps are written right after Phase 1, BEFORE any AI call, so a later
REM Phase 2+ error does NOT affect the parity result -- ignore it if it happens.
set "AI_BACKEND=claude_code"
set "ANTHROPIC_API_KEY="

echo.
echo [%DATE% %TIME%] ===== CAPTURE 1/2: MAIL_BACKEND=com =====
set "MAIL_BACKEND=com"
set "WI_MAIL_PARALLEL=1"
python "%PROJECT_DIR%\fetch_inbox.py"
echo [%DATE% %TIME%] com capture exit %ERRORLEVEL% (non-zero from a later phase is OK - see note above)

echo.
echo [%DATE% %TIME%] ===== CAPTURE 2/2: MAIL_BACKEND=imap =====
set "MAIL_BACKEND=imap"
set "WI_MAIL_PARALLEL=1"
python "%PROJECT_DIR%\fetch_inbox.py"
echo [%DATE% %TIME%] imap capture exit %ERRORLEVEL%

echo.
echo [%DATE% %TIME%] ===== DIFF =====
set "MAIL_BACKEND="
set "WI_MAIL_PARALLEL="
python "%PROJECT_DIR%\diff_mail_pull.py"
set "DIFF_RC=%ERRORLEVEL%"

echo.
echo [%DATE% %TIME%] Parity output: %PROJECT_DIR%\data\parallel\
echo [%DATE% %TIME%] diff exit %DIFF_RC%  (0 = full parity, 1 = differences found - read the report above)
echo.
pause
endlocal
exit /b %DIFF_RC%

:pullfail
echo [%DATE% %TIME%] ERROR: could not download scripts from GitHub. Check network / raw.githubusercontent.com.
pause
endlocal
exit /b 1
