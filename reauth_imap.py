r"""
reauth_imap.py -- one-time / on-demand device-code sign-in for the work-inbox
IMAP+OAuth2 mail backend.

Kevin runs this:
  - once, to prime the token cache before MAIL_BACKEND=imap is ever used, and
  - again whenever the "Outlook mail sign-in expired" toast fires (the silent
    refresh in imap_mail.py could not renew the cached token -- expected
    periodically because this device has no Primary Refresh Token; see the
    no-PRT confirmed-fact memory).

How to run (PowerShell 5.1, which is all Kevin has):
    cd "C:\path\to\work-inbox"
    python .\reauth_imap.py
or just double-click  "Re-auth Work Inbox IMAP.bat".

It prints a short code and a URL, waits for Kevin to approve in a browser
(any device), then writes the refreshed token to
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
    print(f"[{_ts()}] work-inbox IMAP re-auth starting")
    if imap_mail.msal is None:
        print(f"[{_ts()}] FATAL: msal is not installed -- `python -m pip install msal`")
        return 2

    cache = imap_mail._load_cache()
    app = imap_mail.msal.PublicClientApplication(
        imap_mail.CLIENT_ID, authority=imap_mail.AUTHORITY, token_cache=cache
    )

    flow = app.initiate_device_flow(scopes=imap_mail.SCOPES)
    if "user_code" not in flow:
        print(f"[{_ts()}] FATAL: could not start device flow: {flow}")
        return 2

    # Use the canonical https://microsoft.com/devicelogin -- the shortlink
    # https://login.microsoft.com/device misbehaved in the 28 Aug spike.
    print()
    print(f"[{_ts()}] To sign in, open:  https://microsoft.com/devicelogin")
    print(f"[{_ts()}] Enter code:        {flow['user_code']}")
    print(f"[{_ts()}] (this window will wait up to {flow.get('expires_in', 900)}s)")
    print()

    result = app.acquire_token_by_device_flow(flow)  # blocks until done/expired
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
