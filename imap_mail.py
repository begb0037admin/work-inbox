r"""
imap_mail.py -- IMAP + OAuth2 mail-pull backend for work-inbox Phase 1.

Status: FIRST IMPLEMENTATION, behind the MAIL_BACKEND=imap flag. NOT yet
parity-verified against the Outlook-COM pull. Do NOT cut the pipeline over to
this until the parallel-run diff (diff_mail_pull.py) has been clean over
several scheduled cycles AND Kevin has given a fresh explicit go-ahead.
See docs/MAIL_BACKEND_MIGRATION_PLAN.md and docs/PHASE1_IMAP_MIGRATION_AUDIT.md.

What this module does:
  - Acquires an Exchange Online OAuth2 access token via MSAL, silently, from a
    token cache Kevin primed once with reauth_imap.py (device-code flow).
  - Opens a READ-ONLY (IMAP EXAMINE) session to outlook.office365.com:993 with
    SASL XOAUTH2.
  - Pulls the top-level INBOX, a VIP sweep, the five named subfolder trees, and
    Sent Items -- mapping every message to the EXACT internal dict shape
    fetch_inbox.py Phase 1 already produces, so nothing downstream changes.

What this module deliberately does NOT do:
  - No calendar. IMAP has no calendar. Phases 3.7/3.8 stay on Outlook COM.
  - No writes of any kind. imaplib only; smtplib is never imported. The MSAL
    token bundle technically also carries SMTP.Send (Thunderbird's public
    client requests the whole mail bundle) -- this module has no code path
    that could ever use it. See the migration plan's credential section.
  - No Outlook Categories (not an IMAP concept -- and Phase 1 never read them;
    audit confirmed zero dependence).

Credential model (see migration plan step 4 for the full rationale):
  - Public client id = Mozilla Thunderbird's (9e5f94bc-...). No client secret
    exists or is stored. No Oxford app registration, no Oxford IT involvement.
  - Token cache: %LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin.
    Never committed (outside the repo tree; also .gitignore'd defensively).
  - On silent-refresh failure this module raises ImapReauthRequired; the caller
    (fetch_inbox.py) fires ONE rate-limited toast and exits cleanly. It never
    blocks on an interactive prompt.

Every entry point logs a timestamp (via the injected `log` callable).
"""

import os
import re
import ssl
import imaplib
import tempfile
import email
import email.header
import email.utils
from datetime import datetime, timezone

try:
    import msal
except ImportError as _e:  # pragma: no cover - msal is installed on the admin box
    msal = None
    _MSAL_IMPORT_ERROR = _e

# --------------------------------------------------------------------------- #
#  Constants -- proven working in the 28 Aug 2026 feasibility spike
#  (docs/IMAP_OAUTH2_SPIKE_20260828.md).
# --------------------------------------------------------------------------- #
CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"          # Mozilla Thunderbird, public
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPES = ["https://outlook.office365.com/IMAP.AccessAsUser.All"]

IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993

_CACHE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", tempfile.gettempdir()), "WorkInboxAI"
)
TOKEN_CACHE_PATH = os.path.join(_CACHE_DIR, "msal_imap_token_cache.bin")

# Header fields we pull per message (cheap; one FETCH round trip).
_HEADER_FIELDS = ("SUBJECT FROM TO CC DATE MESSAGE-ID IMPORTANCE X-PRIORITY "
                  "X-MSMAIL-PRIORITY CONTENT-CLASS LIST-UNSUBSCRIBE LIST-ID PRECEDENCE")

# Exchange Online IMAP hierarchy separator is "/".
_SEP = "/"


class ImapReauthRequired(Exception):
    """Silent token refresh failed. The caller must toast + exit cleanly,
    never prompt. Carries a short human reason for the log/toast."""


# --------------------------------------------------------------------------- #
#  Token acquisition
# --------------------------------------------------------------------------- #
def _load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_PATH):
        try:
            with open(TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except Exception:
            # A corrupt cache is a reauth condition, not a crash.
            pass
    return cache


def _save_cache(cache):
    if not cache.has_state_changed:
        return
    try:
        os.makedirs(_CACHE_DIR, exist_ok=True)
        with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
            f.write(cache.serialize())
    except Exception:
        pass  # non-fatal; next run just does another silent refresh


def _broker_app(cache):
    """A broker-enabled PublicClientApplication, or None when pymsalruntime
    (msal[broker]) is not importable on this host.

    Proven on the Oxford laptop 2026-08-29 (docs/LAPTOP_MIGRATION_PLAN.md
    Phase 2(i)): the broker CANNOT do the *interactive* acquisition with the
    Thunderbird client id (`broker_error / Status_ApiContractViolation`, twice)
    -- reauth_imap.py does that via the system browser instead -- BUT the
    broker CAN serve a *silent* token from the file cache that browser sign-in
    seeds, and a cold scheduled run resolves its token exactly this way. On the
    admin desktop (no msal[broker]) this returns None and behaviour is
    unchanged (plain app only)."""
    try:
        import pymsalruntime  # noqa: F401
    except Exception:
        return None
    try:
        return msal.PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY, token_cache=cache,
            enable_broker_on_windows=True,
        )
    except Exception:
        return None


def acquire_token_silent(log=print):
    """Return (access_token, upn). Raise ImapReauthRequired on any failure that
    a human sign-in (reauth_imap.py -> system browser) would fix.

    Tries, in the order proven on the laptop: broker-app silent, then plain-app
    silent, both off the shared file cache at TOKEN_CACHE_PATH. Neither prompts.
    """
    if msal is None:
        raise ImapReauthRequired(f"msal not importable: {_MSAL_IMPORT_ERROR!r}")

    cache = _load_cache()
    apps = []
    _b = _broker_app(cache)
    if _b is not None:
        apps.append(("broker", _b))
    apps.append(("plain", msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=cache)))

    errors = []
    for label, app in apps:
        accounts = app.get_accounts()
        if not accounts:
            errors.append(f"{label}: no cached account")
            continue
        account = accounts[0]
        try:
            result = app.acquire_token_silent(SCOPES, account=account)
        except Exception as e:  # pragma: no cover - broker/runtime edge
            errors.append(f"{label}: silent raised {e!r}")
            continue
        if result and "access_token" in result:
            _save_cache(cache)
            upn = account.get("username") or result.get(
                "id_token_claims", {}).get("preferred_username", "")
            log(f"IMAP - silent OAuth2 token OK for {upn} via {label}-app ({_now()})")
            return result["access_token"], upn
        # Common: invalid_grant / AADSTS50173 (token revoked, password change)
        # / AADSTS700082 (refresh token expired past its rolling window).
        err = (result or {}).get("error", "unknown")
        desc = (result or {}).get("error_description", "")
        errors.append(f"{label}: {err} {desc[:120]}".strip())

    _save_cache(cache)
    raise ImapReauthRequired(
        "silent token refresh failed [" + " | ".join(errors)
        + "] -- run 'Re-auth Work Inbox IMAP' once to sign in"
    )


# --------------------------------------------------------------------------- #
#  IMAP session
# --------------------------------------------------------------------------- #
def _imap_connect(access_token, upn, log=print):
    ctx = ssl.create_default_context()
    M = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx)
    auth_bytes = f"user={upn}\x01auth=Bearer {access_token}\x01\x01".encode("utf-8")
    try:
        M.authenticate("XOAUTH2", lambda _challenge: auth_bytes)
    except imaplib.IMAP4.error as e:
        # A rejected bearer here is also a reauth condition.
        raise ImapReauthRequired(f"IMAP XOAUTH2 rejected: {e}")
    log(f"IMAP - AUTHENTICATE OK ({_now()})")
    return M


# --------------------------------------------------------------------------- #
#  Header / body helpers
# --------------------------------------------------------------------------- #
def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _decode_hdr(raw):
    if raw is None:
        return ""
    try:
        parts = email.header.decode_header(raw)
        out = []
        for txt, enc in parts:
            if isinstance(txt, bytes):
                out.append(txt.decode(enc or "utf-8", errors="replace"))
            else:
                out.append(txt)
        # Collapse header folding (CRLF+WSP) and any internal whitespace runs
        # to single spaces -- COM's msg.Subject returns clean single-line text,
        # so the parity diff must too.
        return re.sub(r"\s+", " ", "".join(out)).strip()
    except Exception:
        return re.sub(r"\s+", " ", str(raw)).strip()


def _parse_from(raw_from):
    """-> (display_name, smtp_address). Address case is PRESERVED to match COM's
    msg.SenderEmailAddress (which keeps the sender's original casing);
    case-insensitive handling is done at every comparison site instead."""
    name, addr = email.utils.parseaddr(_decode_hdr(raw_from))
    return name.strip(), (addr or "").strip()


def _importance_from_headers(msg):
    """Map MIME priority headers to Outlook's PR_IMPORTANCE scale: 0 low / 1
    normal / 2 high -- the same mapping Outlook itself uses, so categorise()'s
    `if imp == 2` stays behaviour-equivalent."""
    imp = (msg.get("Importance") or "").strip().lower()
    if imp == "high":
        return 2
    if imp == "low":
        return 0
    xms = (msg.get("X-MSMail-Priority") or "").strip().lower()
    if xms.startswith("high"):
        return 2
    if xms.startswith("low"):
        return 0
    xp = (msg.get("X-Priority") or "").strip()
    m = re.match(r"\s*(\d)", xp)
    if m:
        d = int(m.group(1))
        if d <= 2:
            return 2
        if d >= 4:
            return 0
    return 1


# Every SMTP address that resolves to Kevin's mailbox. kevin.lelitte@... is the
# primary; begb0037@ox.ac.uk is the alias a lot of external/marketing mail is
# addressed to. COM's PropertyAccessor GAL-resolves the alias to the primary,
# so it counts begb0037@ as "Kevin"; the raw IMAP To: header does not -- hence
# this explicit alias set. Extended at runtime with the authenticated UPN.
_KEVIN_ADDRS = {"kevin.lelitte@admin.ox.ac.uk", "begb0037@ox.ac.uk"}


def _kevin_is_primary_recipient(msg, kevin_email):
    """True if any address that resolves to Kevin's mailbox is in To. Over IMAP
    the raw RFC822 To: carries real SMTP addresses (no GAL display-name
    substitution). Fails OPEN (True) on parse failure / no To: header -- same
    philosophy as the COM version: never silently suppress a real email."""
    addrs = {a.lower() for a in _KEVIN_ADDRS} | {(kevin_email or "").lower()}
    try:
        raw_to = msg.get("To", "") or ""
        if not raw_to.strip():
            return True  # no To header at all -> can't tell -> fail open, like COM
        tos = email.utils.getaddresses([raw_to])
        for _n, a in tos:
            if a and a.lower().strip() in addrs:
                return True
        # A parseable To: that doesn't contain Kevin -> he's not the direct
        # recipient. (COM can't always resolve this and fails OPEN; over IMAP
        # the raw To: is authoritative, so a small number of marketing-blast
        # rows will legitimately read False here where COM read True -- the
        # parity report surfaces those for eyeballing rather than us copying
        # COM's fail-open quirk.)
        return False
    except Exception:
        return True


def _has_attachments(full_msg):
    """True if the parsed message carries a real attachment -- any part with a
    filename, an explicit attachment disposition, or a non-text/non-multipart/
    non-message content type. Approximates COM's msg.Attachments.Count > 0
    (which also counts inline images); residual diffs on inline-only mail are
    surfaced by diff_mail_pull.py."""
    if full_msg is None:
        return False
    try:
        if not full_msg.is_multipart():
            return False
        for part in full_msg.walk():
            if part.is_multipart():
                continue
            if part.get_filename():
                return True
            disp = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disp:
                return True
            ct = (part.get_content_type() or "").lower()
            maintype = ct.split("/", 1)[0]
            if maintype not in ("text", "multipart", "message"):
                return True
        return False
    except Exception:
        return False


def _text_preview(raw_message_bytes, limit=150):
    def _decode_part(part):
        raw = part.get_payload(decode=True) or b""
        return raw.decode(part.get_content_charset() or "utf-8", errors="replace")

    try:
        m = email.message_from_bytes(raw_message_bytes)
        plain = ""
        html_txt = ""
        if m.is_multipart():
            for part in m.walk():
                ct = part.get_content_type()
                if part.get_filename():
                    continue
                if ct == "text/plain" and not plain:
                    plain = _decode_part(part)
                elif ct == "text/html" and not html_txt:
                    html_txt = _decode_part(part)
        else:
            if m.get_content_type() == "text/html":
                html_txt = _decode_part(m)
            else:
                plain = _decode_part(m)
        body = plain
        if not body.strip() and html_txt:
            # crude tag strip -- COM's msg.Body returns plain text even for
            # HTML-only mail, so the preview must too.
            body = re.sub(r"(?is)<(script|style).*?</\1>", " ", html_txt)
            body = re.sub(r"(?s)<[^>]+>", " ", body)
            body = re.sub(r"&nbsp;?", " ", body)
            body = re.sub(r"&amp;", "&", body)
        body = re.sub(r"\s+", " ", body).strip()
        return body[:limit]
    except Exception:
        return ""


def _owa_search_link(message_id):
    """OWA deep-link keyed on the internet Message-ID. There is no IMAP path to
    mint an OWA ItemID, so we open OWA's search UI scoped to the exact
    Message-ID. Estate precedent: command-centre sourceType=codex-graph opens a
    stored web_link on outlook.office.com. The dashboard validates this exactly
    like command-centre openEmailWeb(): https + exact-hostname allowlist."""
    if not message_id:
        return ""
    mid = message_id.strip().strip("<>")
    from urllib.parse import quote
    return f"https://outlook.office.com/mail/search?query={quote(mid)}"


# --------------------------------------------------------------------------- #
#  Core per-folder scan
# --------------------------------------------------------------------------- #
def _uid_search_since(M, cutoff):
    """UIDs of messages with INTERNALDATE on/after cutoff's date. IMAP SINCE is
    date-granular; the exact 7-day/6h-grace boundary is re-applied client-side
    against the parsed Date header, matching restrict_date()'s effective
    behaviour."""
    since = cutoff.strftime("%d-%b-%Y")
    typ, data = M.uid("SEARCH", None, "SINCE", since)
    if typ != "OK" or not data or not data[0]:
        return []
    return data[0].split()


def _fetch_one(M, uid, want_body=True):
    """-> (headers_msg, flags_bytes, internaldate_dt, raw_full_bytes_or_None).

    Two FETCH calls: one for FLAGS/INTERNALDATE/headers, one for BODY.PEEK[]
    (the full raw message). Combining multiple BODY.PEEK[...] literals in one
    FETCH makes imaplib's response hard to parse reliably -- that is why
    body_preview came back empty on the first live run. raw_full is always
    fetched (attachment detection needs it, not just the preview)."""
    typ, data = M.uid(
        "FETCH", uid,
        f"(FLAGS INTERNALDATE BODY.PEEK[HEADER.FIELDS ({_HEADER_FIELDS})])",
    )
    if typ != "OK" or not data:
        return None
    header_blob = b""
    flags = b""
    internaldate = None
    for part in data:
        meta = part[0] if isinstance(part, tuple) else part
        if isinstance(part, tuple) and b"HEADER" in (part[0] or b""):
            header_blob = part[1]
        if b"FLAGS" in meta:
            fm = re.search(rb"FLAGS \(([^)]*)\)", meta)
            if fm:
                flags = fm.group(1)
        im = re.search(rb'INTERNALDATE "([^"]+)"', meta)
        if im:
            internaldate = im.group(1).decode("ascii", "replace")
    if not header_blob:
        return None
    msg = email.message_from_bytes(header_blob)
    idt = _parse_internaldate(internaldate)
    if idt is None:
        try:
            idt = email.utils.parsedate_to_datetime(msg.get("Date"))
        except Exception:
            idt = None

    raw_full = None
    try:
        t2, d2 = M.uid("FETCH", uid, "(BODY.PEEK[])")
        if t2 == "OK":
            for part in d2:
                if isinstance(part, tuple) and part[1]:
                    raw_full = part[1]
                    break
    except Exception:
        raw_full = None
    return msg, flags, idt, raw_full


def _parse_internaldate(s):
    """Parse an IMAP INTERNALDATE value ('28-Aug-2026 16:35:29 +0100') into an
    aware datetime. imaplib.Internaldate2tuple wants a whole response line, so
    we hand it a minimal synthetic one; it returns a local-time struct_time
    (epoch-based), which we convert back to an aware datetime."""
    if not s:
        return None
    try:
        import time as _t
        tt = imaplib.Internaldate2tuple(b'x INTERNALDATE "' + s.encode("ascii") + b'"')
        if tt is None:
            return None
        epoch = _t.mktime(tt)
        return datetime.fromtimestamp(epoch).astimezone()
    except Exception:
        # last resort: strptime the common form, assume the given offset
        try:
            return datetime.strptime(s, "%d-%b-%Y %H:%M:%S %z")
        except Exception:
            return None


def _received_str(dt_obj):
    """Match the shape str(msg.ReceivedTime) produces on the COM side
    ('YYYY-MM-DD HH:MM:SS[+TZ]'). Downstream parsing already strips the TZ."""
    if dt_obj is None:
        return ""
    if dt_obj.tzinfo is not None:
        dt_obj = dt_obj.astimezone().replace(tzinfo=None)
    return dt_obj.strftime("%Y-%m-%d %H:%M:%S")


def _build_entry(msg, flags, idt, raw_full, kevin_email, source_folder=None):
    is_read = b"\\Seen" in (flags or b"")
    name, addr = _parse_from(msg.get("From"))
    full = None
    if raw_full:
        try:
            full = email.message_from_bytes(raw_full)
        except Exception:
            full = None
    entry = {
        "subject": _decode_hdr(msg.get("Subject")),
        "from": name or addr,
        "from_email": addr,
        "received": _received_str(idt),
        "is_read": is_read,
        "has_attachments": _has_attachments(full),
        "importance": _importance_from_headers(msg),
        "entry_id": "",  # no IMAP equivalent -- see migration plan
        "message_id": (msg.get("Message-ID") or "").strip(),
        "web_link": _owa_search_link(msg.get("Message-ID") or ""),
        "mail_backend": "imap",  # dashboard opener discriminator
        "kevin_is_primary_recipient": _kevin_is_primary_recipient(msg, kevin_email),
    }
    if source_folder:
        entry["source_folder"] = source_folder
    if not is_read:
        entry["body_preview"] = _text_preview(raw_full or b"")
    return entry


def _scan_mailbox(M, mailbox, cutoff, cap_unread, cap_read, kevin_email,
                  captured_ids, counters, source_folder=None, log=print):
    """Append matching messages from `mailbox` into counters['inbox'].
    `captured_ids` is a set of message_ids already taken (dedup across sweeps)."""
    try:
        typ, _ = M.select(_imap_quote(mailbox), readonly=True)
        if typ != "OK":
            log(f"IMAP - could not EXAMINE {mailbox!r} - skipping ({_now()})")
            return
    except Exception as e:
        log(f"IMAP - EXAMINE {mailbox!r} failed: {e} ({_now()})")
        return

    uids = _uid_search_since(M, cutoff)
    # newest first
    for uid in sorted(uids, key=lambda b: int(b), reverse=True):
        if counters["unread"] >= cap_unread and counters["read"] >= cap_read:
            break
        try:
            fetched = _fetch_one(M, uid)
            if not fetched:
                continue
            msg, flags, idt, raw_full = fetched
            # client-side re-apply of the exact cutoff (SINCE is date-only)
            if idt is not None:
                naive = idt.replace(tzinfo=None) if idt.tzinfo else idt
                if naive < cutoff:
                    continue
            # COM's inbox pull is effectively mail-only (it throws + skips on
            # IPM.Schedule.* items). Match that: drop meeting responses so the
            # briefing doesn't gain "Accepted:/Declined:" rows it never had.
            if "calendarmessage" in (msg.get("Content-Class") or "").lower():
                continue
            if re.match(r"^\s*(accepted|declined|tentative|tentatively accepted|"
                        r"not accepted):\s", _decode_hdr(msg.get("Subject")), re.I):
                continue
            mid = (msg.get("Message-ID") or "").strip()
            if mid and mid in captured_ids:
                continue
            is_read = b"\\Seen" in (flags or b"")
            if is_read and counters["read"] >= cap_read:
                continue
            if not is_read and counters["unread"] >= cap_unread:
                continue
            entry = _build_entry(msg, flags, idt, raw_full,
                                 kevin_email, source_folder=source_folder)
            counters["inbox"].append(entry)
            if mid:
                captured_ids.add(mid)
            if is_read:
                counters["read"] += 1
            else:
                counters["unread"] += 1
        except Exception:
            continue  # per-message failure never fails the pull (COM parity)


def _imap_quote(name):
    return '"' + name.replace('"', '\\"') + '"'


def _mutf7_encode_ascii(name):
    """Minimal IMAP modified-UTF-7 for the ASCII folder names we use: only '&'
    needs escaping ('&' -> '&-'). Exchange Online returns e.g. 'INBOX/H&-S' for
    a folder named 'H&S'. None of the SUBFOLDER_TREES contain non-ASCII, so the
    full codec isn't needed here."""
    return name.replace("&", "&-")


_LIST_RE = re.compile(rb'^\((?P<flags>[^)]*)\)\s+(?P<sep>"[^"]*"|NIL)\s+(?P<name>.+)$')


def _list_mailboxes(M):
    """-> [(name_str, flags_str), ...] over the whole account. name_str is left
    in the server's modified-UTF-7 form (compare against _mutf7_encode_ascii)."""
    typ, data = M.list()
    out = []
    if typ != "OK" or not data:
        return out
    for row in data:
        if not row:
            continue
        m = _LIST_RE.match(row.strip())
        if not m:
            continue
        name = m.group("name").decode("ascii", "replace").strip()
        if name.startswith('"') and name.endswith('"'):
            name = name[1:-1]
        flags = m.group("flags").decode("ascii", "replace")
        out.append((name, flags))
    return out


# --------------------------------------------------------------------------- #
#  Public entry point
# --------------------------------------------------------------------------- #
def pull(cutoff, *, kevin_email, vip_names, vip_emails, subfolder_trees,
         max_unread=50, max_read=30, sub_max_unread=40, sub_max_read=20,
         log=print):
    """The IMAP equivalent of fetch_inbox.py Phase 1's mail pull.

    Returns {"inbox": [...], "sent": [...], "meta": {...}} where every dict has
    the exact keys the COM path produces (plus message_id / web_link /
    mail_backend, which downstream code ignores unless it opts in).

    Raises ImapReauthRequired if the cached token cannot be silently refreshed.
    """
    log(f"IMAP mail pull starting ({_now()}) - cutoff {cutoff:%Y-%m-%d %H:%M}")
    token, upn = acquire_token_silent(log=log)
    if upn:
        _KEVIN_ADDRS.add(upn.lower())
    M = _imap_connect(token, upn, log=log)
    try:
        captured_ids = set()

        # --- top-level INBOX ---
        counters = {"inbox": [], "unread": 0, "read": 0}
        _scan_mailbox(M, "INBOX", cutoff, max_unread, max_read, kevin_email,
                      captured_ids, counters, source_folder=None, log=log)
        inbox = counters["inbox"]
        inbox.sort(key=lambda x: (not x["is_read"], x["received"]), reverse=True)
        log(f"IMAP - INBOX: {len(inbox)} (unread {counters['unread']} / read {counters['read']})")

        # --- VIP sweep (uncapped, whole 7d window, dedup on message_id) ---
        vip_names_l = {n.strip() for n in vip_names}
        vip_emails_l = {e.lower().strip() for e in vip_emails}
        vip_counter = {"inbox": [], "unread": 0, "read": 0}
        _scan_mailbox(M, "INBOX", cutoff, 10**9, 10**9, kevin_email,
                      set(captured_ids), vip_counter, source_folder=None, log=log)
        vip_added = 0
        for e in vip_counter["inbox"]:
            if e["message_id"] and e["message_id"] in captured_ids:
                continue
            if (e["from"] or "").strip() in vip_names_l or (e["from_email"] or "").lower() in vip_emails_l:
                inbox.append(e)
                if e["message_id"]:
                    captured_ids.add(e["message_id"])
                vip_added += 1
        inbox.sort(key=lambda x: (not x["is_read"], x["received"]), reverse=True)
        log(f"IMAP - VIP sweep added {vip_added} - total {len(inbox)}")

        # --- named subfolder trees ---
        all_boxes = [n for (n, _f) in _list_mailboxes(M)]
        sub_counter = {"inbox": [], "unread": 0, "read": 0}
        sub_added_before = len(inbox)
        for tree in subfolder_trees:
            target = "INBOX" + _SEP + _mutf7_encode_ascii(tree)
            matches = [b for b in all_boxes
                       if b == target or b.startswith(target + _SEP)]
            if not matches:
                log(f"IMAP - no subfolder matching {target!r} "
                    f"(tree {tree!r}) - skipped this run")
                continue
            for box in matches:
                if (sub_counter["unread"] >= sub_max_unread
                        and sub_counter["read"] >= sub_max_read):
                    break
                _scan_mailbox(M, box, cutoff, sub_max_unread, sub_max_read,
                              kevin_email, captured_ids, sub_counter,
                              source_folder=box, log=log)
        for e in sub_counter["inbox"]:
            inbox.append(e)
        inbox.sort(key=lambda x: (not x["is_read"], x["received"]), reverse=True)
        log(f"IMAP - subfolder sweep added {len(inbox) - sub_added_before} "
            f"(unread {sub_counter['unread']} / read {sub_counter['read']})")

        # --- Sent Items ---
        sent = _pull_sent(M, cutoff, log=log)
        log(f"IMAP - Sent Items: {len(sent)}")

        unread_total = sum(1 for m in inbox if not m["is_read"])
        meta = {
            "upn": upn,
            "inbox_total": len(inbox),
            "inbox_unread": unread_total,
            "sent_total": len(sent),
            "pulled_at": _now(),
        }
        log(f"IMAP mail pull done ({_now()}) - inbox {len(inbox)} "
            f"(unread {unread_total}) sent {len(sent)}")
        return {"inbox": inbox, "sent": sent, "meta": meta}
    finally:
        try:
            M.logout()
        except Exception:
            pass


def _pull_sent(M, cutoff, log=print):
    # Prefer the \Sent special-use box; fall back to the well-known name.
    target = None
    typ, data = M.list()
    if typ == "OK" and data:
        for row in data:
            s = (row or b"").decode("utf-8", errors="replace")
            if "\\Sent" in s:
                m = re.search(r'"?([^"]*)"?$', s.strip())
                if m:
                    target = m.group(1)
                    break
    for candidate in ([target] if target else []) + ['Sent Items', 'Sent']:
        if not candidate:
            continue
        try:
            typ, _ = M.select(_imap_quote(candidate), readonly=True)
            if typ == "OK":
                target = candidate
                break
        except Exception:
            continue
    else:
        log("IMAP - no Sent mailbox found - sent sweep skipped")
        return []

    out = []
    seen_mids = set()
    _resp_prefix = re.compile(
        r"^\s*(accepted|declined|tentative|tentatively accepted|not accepted|"
        r"cancelled|canceled):\s", re.I)
    for uid in sorted(_uid_search_since(M, cutoff), key=lambda b: int(b), reverse=True):
        try:
            fetched = _fetch_one(M, uid)
            if not fetched:
                continue
            msg, flags, idt, raw_full = fetched
            if idt is not None:
                naive = idt.replace(tzinfo=None) if idt.tzinfo else idt
                if naive < cutoff:
                    continue
            subject = _decode_hdr(msg.get("Subject"))
            # Exclude meeting requests / responses / calendar-invite copies that
            # land in Sent Items -- the COM Sent pull never surfaces these (its
            # dt(msg.SentOn) filter drops IPM.Schedule.* items). Match COM's
            # clean set: Content-Class header (covers requests AND responses),
            # subject prefix, or a text/calendar part.
            if "calendarmessage" in (msg.get("Content-Class") or "").lower():
                continue
            if _resp_prefix.match(subject or ""):
                continue
            full = None
            if raw_full:
                try:
                    full = email.message_from_bytes(raw_full)
                except Exception:
                    full = None
            if full is not None:
                if "calendar" in (full.get_content_type() or ""):
                    continue
                if full.is_multipart() and any(
                    p.get_content_type() == "text/calendar" for p in full.walk()
                ):
                    continue
            # Exchange strips the iCal part from the SENT copy of a meeting
            # request / calendar-share, leaving a plain text/html item that has
            # no header or MIME tell -- but COM excludes it via MessageClass
            # (IPM.Schedule.*). The one reliable remaining signal on this
            # mailbox is the invite body preamble.
            _bt = _text_preview(raw_full or b"", limit=400).lower()
            if _bt.startswith("microsoft outlook web access:"):
                continue
            if ("\nwhen:" in _bt or " when:" in _bt[:80]) and \
               ("\nwhere:" in _bt or "\nlocation:" in _bt or "outlook web access" in _bt):
                continue
            mid = (msg.get("Message-ID") or "").strip()
            mid_norm = mid.strip("<>").lower()
            if mid_norm and mid_norm in seen_mids:
                continue
            if mid_norm:
                seen_mids.add(mid_norm)
            out.append({
                "subject": subject,
                "to": _decode_hdr(msg.get("To")),
                "sent": _received_str(idt),
                "body_preview": _text_preview(raw_full or b"", limit=100),
                "entry_id": "",
                "message_id": mid,
            })
        except Exception:
            continue
    return out


# --------------------------------------------------------------------------- #
#  Resolution-tracking helper (Phase 3.9) -- IMAP replacement for
#  mapi.GetItemFromID(eid).Parent.EntryID == inbox.EntryID
# --------------------------------------------------------------------------- #
def message_still_in_inbox(message_ids, log=print):
    """Given an iterable of internet Message-IDs, return the subset that is
    still present in INBOX right now. A tracked id that is NOT returned has
    been filed/deleted -> 'moved_out' -> resolved. A SEARCH failure for an id
    means it is simply omitted from 'still present', which is the fail-OPEN
    direction only if the caller treats 'absent' as 'carry' -- BUT Phase 3.9
    wants absent == resolved, so callers should treat a raised exception here
    as 'inconclusive, carry all' rather than trusting a partial result."""
    token, upn = acquire_token_silent(log=log)
    M = _imap_connect(token, upn, log=log)
    still = set()
    try:
        M.select('"INBOX"', readonly=True)
        for mid in message_ids:
            if not mid:
                continue
            clean = mid.strip()
            try:
                typ, data = M.uid("SEARCH", None, "HEADER", "Message-ID", clean)
                if typ == "OK" and data and data[0].strip():
                    still.add(mid)
            except Exception:
                # Unknown for this id -- caller decides. Record nothing.
                continue
        return still
    finally:
        try:
            M.logout()
        except Exception:
            pass
