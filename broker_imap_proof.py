r"""
broker_imap_proof.py -- work-inbox laptop migration, Phase 2(i) MAKE-OR-BREAK.

Proves that MSAL's Windows broker (WAM) can acquire an Exchange Online IMAP
access token SILENTLY off this laptop's Primary Refresh Token, and that a cold
second run of this script needs NO prompt at all.

READ-ONLY. The only thing it writes is its own MSAL token cache at
  %LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin
It does an IMAP EXAMINE (read-only) of INBOX and prints the message count and one
header. It never sends, moves, marks, or deletes anything.

HOW TO RUN  (on the laptop, signed in as ad-oak\begb0037):
    python broker_imap_proof.py        # run 1 -- approve the sign-in IF a dialog appears
    python broker_imap_proof.py        # run 2, COLD -- must reach "=== PASS ===" with NO prompt

Exit codes:  0 = PASS (silent token + IMAP OK)   1 = PARTIAL (interactive needed this
run -- run again)   3 = FAIL (prompt on a run that should have been silent, no token,
or IMAP auth failed)
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


def main():
    log("=== broker_imap_proof START ===")
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
        log("pymsalruntime import OK -- broker/WAM will be used (enable_broker_on_windows=True)")
    except Exception as e:
        log(f"WARNING: pymsalruntime NOT importable ({e!r}) -- broker cannot be used; this run "
            f"cannot prove the broker path")

    os.makedirs(CACHE_DIR, exist_ok=True)
    cache = msal.SerializableTokenCache()
    first_run_hint = True
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "r", encoding="utf-8") as fh:
                cache.deserialize(fh.read())
            log(f"loaded existing MSAL file cache ({os.path.getsize(CACHE_PATH)} bytes) "
                f"-- NOT a first run")
            first_run_hint = False
        except Exception as e:
            log(f"could not read existing cache ({e!r}) -- treating as first run")
    else:
        log("no existing MSAL file cache on disk -- looks like a first run")

    try:
        app = msal.PublicClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            token_cache=cache,
            enable_broker_on_windows=broker_ok,
        )
    except Exception as e:
        log(f"FAIL: PublicClientApplication construction failed: {e!r}")
        return 3

    def persist():
        if cache.has_state_changed:
            try:
                with open(CACHE_PATH, "w", encoding="utf-8") as fh:
                    fh.write(cache.serialize())
                log(f"token cache persisted ({os.path.getsize(CACHE_PATH)} bytes)")
            except Exception as e:
                log(f"WARNING: could not persist token cache: {e!r}")

    # ---------------------------------------------------------------- #
    # 1. SILENT first -- this is the property we are trying to prove.
    # ---------------------------------------------------------------- #
    result = None
    accounts = app.get_accounts()
    log(f"app.get_accounts() -> {len(accounts)} account(s): "
        f"{[a.get('username') for a in accounts]}")
    if accounts:
        log("attempting acquire_token_silent (NO prompt expected) ...")
        try:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
        except Exception as e:
            log(f"acquire_token_silent raised: {e!r}")
            result = None
        if result and "access_token" in result:
            log("SILENT token acquisition OK  <-- no prompt, this is the pass path")
        else:
            err = (result or {}).get("error")
            log(f"silent returned no token (error={err}) -- will fall back to interactive")
            result = None
    else:
        log("no cached/broker account visible yet -- interactive is expected on THIS run only")

    silent_worked = result is not None

    # ---------------------------------------------------------------- #
    # 2. INTERACTIVE fallback -- only runs if silent could not.
    # ---------------------------------------------------------------- #
    if not result:
        if not broker_ok:
            log("FAIL: no broker and no silent token -- cannot prove anything without pymsalruntime")
            return 3
        log("attempting acquire_token_interactive via broker/WAM ...")
        log("   >>> IF A SIGN-IN DIALOG APPEARS, NOTE EXACTLY WHAT IT ASKS FOR:")
        log("       (a) just pick the ad-oak\\begb0037 account   (b) a consent/'let this app' screen")
        log("       (c) a full password + MFA prompt   -- (c) would be a concern")
        try:
            phc = getattr(msal.PublicClientApplication, "CONSOLE_WINDOW_HANDLE", None)
            kwargs = {"parent_window_handle": phc} if phc is not None else {}
            t0 = datetime.now(timezone.utc)
            result = app.acquire_token_interactive(SCOPES, **kwargs)
            secs = (datetime.now(timezone.utc) - t0).total_seconds()
            log(f"interactive call returned after {secs:.1f}s")
        except Exception as e:
            log(f"FAIL: acquire_token_interactive raised: {e!r}")
            persist()
            return 3
        if not result or "access_token" not in result:
            log(f"FAIL: interactive got no token: error={(result or {}).get('error')} "
                f"desc={str((result or {}).get('error_description', ''))[:300]}")
            persist()
            return 3
        log("interactive token acquisition OK")

    persist()

    # ---------------------------------------------------------------- #
    # 2b. Immediate SECOND silent call in the same process.
    # ---------------------------------------------------------------- #
    accounts = app.get_accounts()
    if accounts:
        log("re-running acquire_token_silent immediately (same process) ...")
        try:
            r2 = app.acquire_token_silent(SCOPES, account=accounts[0])
        except Exception as e:
            r2 = None
            log(f"second silent raised: {e!r}")
        if r2 and "access_token" in r2:
            log("second silent OK -- token is now cached for the next cold run")
        else:
            log(f"second silent returned no token (error={(r2 or {}).get('error')})")

    token = result["access_token"]
    claims = result.get("id_token_claims") or {}
    upn = claims.get("preferred_username") or claims.get("upn") or FALLBACK_UPN
    log(f"token OK. upn={upn}  granted_scopes={result.get('scope')}")

    # ---------------------------------------------------------------- #
    # 3. IMAP read-only proof.
    # ---------------------------------------------------------------- #
    imap_ok = False
    try:
        log(f"connecting {IMAP_HOST}:{IMAP_PORT} (TLS) ...")
        ctx = ssl.create_default_context()
        M = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx)
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

    # ---------------------------------------------------------------- #
    # Verdict.
    # ---------------------------------------------------------------- #
    log("-------------------------------------------------------------")
    log(f"broker_used             : {broker_ok}")
    log(f"first_run_hint          : {first_run_hint}  (no prior file cache)")
    log(f"silent_token_no_prompt  : {silent_worked}")
    log(f"imap_examine_inbox_ok   : {imap_ok}")
    if silent_worked and imap_ok:
        log("=== PASS === broker gave a token with NO prompt AND IMAP read works.")
        log("            Run this script ONCE MORE (cold) to double-confirm the silent path.")
        return 0
    if (not silent_worked) and imap_ok:
        log("=== PARTIAL === interactive was needed on THIS run (expected on the very first run).")
        log("            RUN THE SCRIPT AGAIN now. The second run MUST print '=== PASS ==='")
        log("            with NO dialog. If the second run prompts, that is a FAIL -- report it.")
        return 1
    log("=== FAIL === see the log above.")
    return 3


if __name__ == "__main__":
    sys.exit(main())
