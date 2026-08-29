r"""
broker_imap_proof.py -- work-inbox laptop migration, Phase 2(i) MAKE-OR-BREAK.  v2.

v1 finding (29 Aug 15:18Z): MSAL broker/WAM interactive returned an instant
`broker_error / Status_ApiContractViolation` with the Thunderbird client id
`9e5f94bc-...` -- that client is not WAM/broker-enabled (device-code + loopback
browser only). We cannot swap the client id: Thunderbird's is the ONLY one
confirmed to get IMAP.AccessAsUser.All at Oxford.

v2 tries auth paths in order and logs which one wins:
  PATH A  broker interactive with a REAL HWND (GetConsoleWindow, else GetDesktopWindow)
  PATH B  plain interactive, NO broker -- system browser, already PRT-SSO'd on this
          laptop, so expect zero-to-one click, no password, no MFA
Then the real proof (either path): a COLD second run must get a token via
`acquire_token_silent` from the file cache with NO prompt.

READ-ONLY. Writes only its own MSAL token cache at
  %LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin
IMAP EXAMINE (read-only) of INBOX only. Never sends / moves / marks / deletes.

HOW TO RUN  (laptop, signed in as ad-oak\begb0037):
    python broker_imap_proof.py        # run 1 -- do the SSO click if the browser opens
    python broker_imap_proof.py        # run 2, COLD -- must reach "=== PASS ===" with NO prompt

GRADING:
  broker path (A) works                              -> FULL WIN (silent forever)
  fallback path (B) works AND run 2 is silent        -> ACCEPTABLE WIN. Day-to-day
     scheduled runs are silent; the periodic (~14-90d) re-auth is one SSO click on
     the laptop, no password. Proceed to Phase 3.
  run 2 still prompts / no token                     -> FAIL. Escalate (device-code
     via reauth_imap.py as last resort).

Exit: 0 = PASS (silent + IMAP OK)   1 = PARTIAL (interactive needed this run -- run
again)   3 = FAIL
"""

import os
import ssl
import sys
import imaplib
from datetime import datetime, timezone

CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"          # Mozilla Thunderbird public client
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPES = ["https://outlook.office365.com/IMAP.AccessAsUser.All"]
IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993
CACHE_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "WorkInboxAI")
CACHE_PATH = os.path.join(CACHE_DIR, "msal_imap_token_cache.bin")
FALLBACK_UPN = "begb0037@ox.ac.uk"


def log(msg):
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%SZ")
    print(f"{ts}  {msg}", flush=True)


def _real_hwnd():
    """A genuine top-level window handle for the broker to parent its dialog to.
    Avoids the 64-bit truncation trap by setting restype to c_void_p."""
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        k32.GetConsoleWindow.restype = ctypes.c_void_p
        h = k32.GetConsoleWindow()
        if not h:
            u32 = ctypes.windll.user32
            u32.GetDesktopWindow.restype = ctypes.c_void_p
            h = u32.GetDesktopWindow()
        return int(h) if h else 0
    except Exception as e:
        log(f"could not resolve a real HWND ({e!r}) -- will pass 0")
        return 0


def _try_silent(app, label):
    try:
        accounts = app.get_accounts()
    except Exception as e:
        log(f"{label}: get_accounts raised {e!r}")
        return None
    log(f"{label}: get_accounts -> {len(accounts)}: {[a.get('username') for a in accounts]}")
    if not accounts:
        return None
    try:
        r = app.acquire_token_silent(SCOPES, account=accounts[0])
    except Exception as e:
        log(f"{label}: acquire_token_silent raised {e!r}")
        return None
    if r and "access_token" in r:
        log(f"{label}: SILENT token OK  <-- no prompt")
        return r
    log(f"{label}: silent returned no token (error={(r or {}).get('error')})")
    return None


def main():
    log("=== broker_imap_proof v2 START ===")
    log(f"python {sys.version.split()[0]}")
    log(f"whoami: {os.environ.get('USERDOMAIN', '?')}\\{os.environ.get('USERNAME', '?')}")
    log(f"token cache: {CACHE_PATH}")

    try:
        import msal
    except Exception as e:
        log(f"FAIL: cannot import msal: {e!r}")
        return 3
    log(f"msal {getattr(msal, '__version__', '?')}")

    broker_ok = False
    try:
        import pymsalruntime  # noqa: F401
        broker_ok = True
        log("pymsalruntime import OK -- PATH A (broker) will be attempted")
    except Exception as e:
        log(f"pymsalruntime NOT importable ({e!r}) -- PATH A skipped, PATH B only")

    os.makedirs(CACHE_DIR, exist_ok=True)
    cache = msal.SerializableTokenCache()
    primed = False
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "r", encoding="utf-8") as fh:
                cache.deserialize(fh.read())
            log(f"loaded existing file cache ({os.path.getsize(CACHE_PATH)} bytes) -- NOT a first run")
            primed = True
        except Exception as e:
            log(f"could not read existing cache ({e!r}) -- treating as first run")
    else:
        log("no existing file cache -- looks like a first run")

    # One shared cache, up to two app objects (broker + plain).
    app_plain = msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=cache, enable_broker_on_windows=False
    )
    app_broker = None
    if broker_ok:
        try:
            app_broker = msal.PublicClientApplication(
                CLIENT_ID, authority=AUTHORITY, token_cache=cache, enable_broker_on_windows=True
            )
        except Exception as e:
            log(f"broker app construction failed ({e!r}) -- PATH A unavailable")
            app_broker = None

    def persist():
        if cache.has_state_changed:
            try:
                with open(CACHE_PATH, "w", encoding="utf-8") as fh:
                    fh.write(cache.serialize())
                log(f"token cache persisted ({os.path.getsize(CACHE_PATH)} bytes)")
            except Exception as e:
                log(f"WARNING: could not persist cache: {e!r}")

    # ------------------------------------------------------------------ #
    # 1. SILENT first (broker app, then plain app) -- the win condition.
    # ------------------------------------------------------------------ #
    result = None
    auth_path = None
    if app_broker is not None:
        result = _try_silent(app_broker, "silent[broker]")
        if result:
            auth_path = "silent(broker-app)"
    if result is None:
        result = _try_silent(app_plain, "silent[plain]")
        if result:
            auth_path = "silent(plain-app)"

    silent_worked = result is not None

    # ------------------------------------------------------------------ #
    # 2. INTERACTIVE -- only if silent could not. PATH A then PATH B.
    # ------------------------------------------------------------------ #
    if result is None:
        # ---- PATH A: broker with a real HWND ----
        if app_broker is not None:
            hwnd = _real_hwnd()
            log(f"PATH A: broker interactive, parent_window_handle={hwnd} ...")
            log("   >>> if a WINDOWS dialog appears, note what it asks for (account pick / consent / password)")
            try:
                r = app_broker.acquire_token_interactive(SCOPES, parent_window_handle=hwnd)
            except Exception as e:
                r = None
                log(f"PATH A raised: {e!r}")
            if r and "access_token" in r:
                result = r
                auth_path = "broker+hwnd"
                log("PATH A OK -- broker interactive token acquired")
            else:
                log(f"PATH A no token: error={(r or {}).get('error')} "
                    f"desc={str((r or {}).get('error_description', ''))[:300]}")
                log("PATH A did not work with this client id -- moving to PATH B (no loop)")

        # ---- PATH B: plain, system browser, no broker ----
        if result is None:
            log("PATH B: plain interactive via system browser (no broker) ...")
            log("   >>> the browser is already PRT-SSO'd on this laptop: expect 0-1 clicks, NO password, NO MFA")
            try:
                r = app_plain.acquire_token_interactive(SCOPES)
            except Exception as e:
                r = None
                log(f"PATH B raised: {e!r}")
            if r and "access_token" in r:
                result = r
                auth_path = "plain-browser"
                log("PATH B OK -- system-browser token acquired")
            else:
                log(f"PATH B no token: error={(r or {}).get('error')} "
                    f"desc={str((r or {}).get('error_description', ''))[:300]}")

    if result is None or "access_token" not in result:
        persist()
        log("=== FAIL === no token from any path. See the log above.")
        return 3

    persist()

    # Immediate second silent (plain app) to confirm the file cache is seeded for run 2.
    r2 = _try_silent(app_plain, "silent[plain, re-check]")
    log(f"file-cache seeded for a cold run 2: {bool(r2)}")

    token = result["access_token"]
    claims = result.get("id_token_claims") or {}
    upn = claims.get("preferred_username") or claims.get("upn") or FALLBACK_UPN
    log(f"AUTH PATH THIS RUN: {auth_path}")
    log(f"token OK. upn={upn}  granted_scopes={result.get('scope')}")

    # ------------------------------------------------------------------ #
    # 3. IMAP read-only proof.
    # ------------------------------------------------------------------ #
    imap_ok = False
    try:
        log(f"connecting {IMAP_HOST}:{IMAP_PORT} (TLS) ...")
        M = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ssl.create_default_context())
        auth_bytes = f"user={upn}\x01auth=Bearer {token}\x01\x01".encode("utf-8")
        typ, data = M.authenticate("XOAUTH2", lambda _=None: auth_bytes)
        log(f"IMAP AUTHENTICATE XOAUTH2 -> {typ} {data}")
        typ, data = M.select("INBOX", readonly=True)   # readonly=True => IMAP EXAMINE
        log(f"EXAMINE INBOX -> {typ} {data}")
        if typ == "OK" and data and data[0] is not None:
            count = int(data[0])
            log(f"INBOX message count: {count}")
            if count > 0:
                typ, hdr = M.fetch(str(count).encode("ascii"),
                                   "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
                if typ == "OK" and hdr and hdr[0]:
                    text = hdr[0][1].decode("utf-8", "replace").strip().replace("\r\n", " | ")
                    log(f"latest message header: {text[:300]}")
            imap_ok = True
        try:
            M.close()
        except Exception:
            pass
        M.logout()
    except Exception as e:
        log(f"FAIL: IMAP step raised: {e!r}")
        return 3

    # ------------------------------------------------------------------ #
    # Verdict.
    # ------------------------------------------------------------------ #
    log("-------------------------------------------------------------")
    log(f"had_primed_cache        : {primed}")
    log(f"auth_path_this_run      : {auth_path}")
    log(f"silent_token_no_prompt  : {silent_worked}")
    log(f"imap_examine_inbox_ok   : {imap_ok}")
    if silent_worked and imap_ok:
        won = "FULL WIN (broker)" if "broker" in (auth_path or "") else "ACCEPTABLE WIN (fallback)"
        log(f"=== PASS === {won}: silent token, NO prompt, IMAP read OK.")
        log("            (Run once more cold if you want a third confirmation.)")
        return 0
    if (not silent_worked) and imap_ok:
        log(f"=== PARTIAL === interactive was needed this run via: {auth_path}")
        log("            RUN THE SCRIPT AGAIN now (cold). The second run MUST print")
        log("            '=== PASS ===' via a silent path with NO browser / NO dialog.")
        return 1
    log("=== FAIL === see the log above.")
    return 3


if __name__ == "__main__":
    sys.exit(main())
