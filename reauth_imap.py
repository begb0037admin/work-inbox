r"""
reauth_imap.py -- one-time / on-demand sign-in that seeds the work-inbox
IMAP+OAuth2 token cache.

Kevin runs this:
  - once, to prime the token cache before MAIL_BACKEND=imap is ever used, and
  - again whenever the "IMAP mail sign-in expired" toast fires (the silent
    refresh in imap_mail.acquire_token_silent could not renew the cached
    token -- expected periodically: Conditional-Access sign-in-frequency /
    ~90d refresh-token roll).

Default = system-browser sign-in. Proven on the Oxford laptop 2026-08-29
(docs/LAPTOP_MIGRATION_PLAN.md Phase 2(i)): the Thunderbird client id is NOT
WAM/broker-interactive-capable, but a plain system-browser sign-in on the
PRT-joined laptop is ONE account click -- no password, no MFA. The scheduled
runs then refresh silently via imap_mail.acquire_token_silent (broker-app
path) from the cache this seeds.

    cd "C:\path\to\work-inbox"
    python .\reauth_imap.py                 # opens a browser, one click
    python .\reauth_imap.py --device-code   # fallback: sign in on another device

Writes the refreshed token to
%LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin and verifies it with a
read-only INBOX check. Nothing is sent anywhere. No secret is stored (the
client id is Mozilla Thunderbird's public one). Prints timestamps.
"""

import sys
import time
from datetime import datetime

import imap_mail  # same directory


def _ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def main():
    device_code = "--device-code" in sys.argv[1:]
    mode = "device-code" if device_code else "system-browser"
    print(f"[{_ts()}] work-inbox IMAP re-auth starting ({mode} mode)")
    if imap_mail.msal is None:
        print(f"[{_ts()}] FATAL: msal is not installed -- "
              f"`python -m pip install \"msal[broker]\"`")
        return 2

    cache = imap_mail._load_cache()
    app = imap_mail.msal.PublicClientApplication(
        imap_mail.CLIENT_ID, authority=imap_mail.AUTHORITY, token_cache=cache
    )

    if device_code:
        flow = app.initiate_device_flow(scopes=imap_mail.SCOPES)
        if "user_code" not in flow:
            print(f"[{_ts()}] FATAL: could not start device flow: {flow}")
            return 2
        # Canonical https://microsoft.com/devicelogin -- the shortlink
        # https://login.microsoft.com/device misbehaved in the 28 Aug spike.
        print()
        print(f"[{_ts()}] To sign in, open:  https://microsoft.com/devicelogin")
        print(f"[{_ts()}] Enter code:        {flow['user_code']}")
        print(f"[{_ts()}] (this window will wait up to {flow.get('expires_in', 900)}s)")
        print()
        result = app.acquire_token_by_device_flow(flow)  # blocks until done/expired
    else:
        # PATH B from the Phase 2(i) proof: plain interactive, system browser,
        # NO broker. On the PRT-joined laptop this is one account click.
        print(f"[{_ts()}] a browser window will open -- pick the ad-oak\\begb0037 "
              f"account (one click; no password expected on the PRT-joined laptop)")
        result = app.acquire_token_interactive(imap_mail.SCOPES)

    imap_mail._save_cache(cache)

    if not result or "access_token" not in result:
        err = (result or {}).get("error", "unknown")
        desc = (result or {}).get("error_description", "")
        print(f"[{_ts()}] FATAL: sign-in did not complete: {err} {desc[:200]}")
        return 1

    granted = result.get("scope", "")
    print(f"[{_ts()}] token acquired. granted scopes: {granted}")

    # Verify with a read-only INBOX check.
    try:
        upn = (app.get_accounts() or [{}])[0].get("username", "")
        M = imap_mail._imap_connect(result["access_token"], upn)
        typ, data = M.select('"INBOX"', readonly=True)
        count = data[0].decode() if (typ == "OK" and data and data[0]) else "?"
        M.logout()
        print(f"[{_ts()}] verified: EXAMINE INBOX OK, {count} messages, as {upn}")
    except Exception as e:
        print(f"[{_ts()}] WARNING: token cached but verification read failed: {e}")
        return 1

    print(f"[{_ts()}] done. Cache written to {imap_mail.TOKEN_CACHE_PATH}")
    print(f"[{_ts()}] The scheduled briefing will now refresh this silently until it next expires.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
