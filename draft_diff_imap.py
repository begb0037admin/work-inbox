r"""
draft_diff_imap.py -- IMAP + OAuth2 backend for draft_final_diff_capture.py.

The MAIL_BACKEND=imap path for the ongoing draft/final diff capture, exactly
mirroring imap_mail.py (which is MAIL_BACKEND=imap for the main Phase 1 mail
pull). Same credential model, same read-only EXAMINE discipline, same
"smtplib is never imported" rule. Behind the MAIL_BACKEND flag:
  MAIL_BACKEND=com (default)  -> draft_final_diff_capture.py stays on Outlook
                                COM, byte-identical, this module never loads.
  MAIL_BACKEND=imap           -> draft_final_diff_capture.py uses this module
                                and never imports win32com.

WHAT THIS REPLACES
  COM path                          -> IMAP equivalent here
  --------------------------------------------------------------------------
  mapi.GetDefaultFolder(16) Drafts  -> EXAMINE of the \Drafts special-use
                                       mailbox ("Drafts" on this account)
  mapi.GetDefaultFolder(5)  Sent    -> EXAMINE of \Sent ("Sent Items")
  msg.Class == 43 (olMail)          -> RFC822 meeting-item heuristics (the
                                       same set imap_mail.py uses: Content-
                                       Class, subject prefix, text/calendar)
  msg.ConversationID                -> conversation GUID decoded from the
                                       Thread-Index header (see _conv_key)
  msg.ConversationTopic             -> Thread-Topic header (else Subject)
  msg.Subject / .To                 -> decoded Subject / To headers
  msg.Body                          -> full text body (_full_body), NOT
                                       truncated -- redaction + AI need it all
  msg.SentOn                        -> INTERNALDATE (else Date header)
  msg.EntryID                       -> NO IMAP EQUIVALENT. The caller records
                                       the internet Message-ID instead, as
                                       provenance only (EntryID was never used
                                       for a lookup -- only written into the
                                       diff's confirmed_via string).

KNOWN FIDELITY GAP -- the correlation key.
  The whole mechanism correlates a vanished draft with its sent counterpart by
  a stable per-conversation key. Over COM that is msg.ConversationID. IMAP does
  not expose PR_CONVERSATION_ID. This module derives an equivalent from the
  Thread-Index header (PR_CONVERSATION_INDEX): its decoded bytes [6:22] are the
  16-byte conversation GUID that Outlook itself hashes to form ConversationID,
  and Exchange stamps it on both the saved draft and the eventual sent item.
  The derived value is NOT byte-equal to Outlook's ConversationID and does not
  need to be -- it only needs to be identical between a draft and its own sent
  version, which it is. Fallbacks when Thread-Index is absent: the thread-root
  Message-ID from References / In-Reply-To, then a normalised Thread-Topic --
  the same "secondary signal only" stance the COM path takes with
  ConversationTopic. Because the key namespace differs from the COM ledger, the
  imap path uses its own ledger/out-dir (the wrapper passes --ledger-path /
  --out-dir); a first run there re-baselines from scratch and produces zero
  pairs, exactly as a first COM run does -- by design, not a bug.

Every entry point logs a timestamp via the injected `log` callable.
"""

import re
import base64
import email
import email.utils
from datetime import datetime, timedelta

# Reuse imap_mail's proven auth + connection + header helpers verbatim. If
# imap_mail's constants (client id, authority, scope, host) ever change, this
# module follows automatically -- one place, no drift.
from imap_mail import (
    acquire_token_silent,
    _imap_connect,
    _imap_quote,
    _parse_internaldate,
    _decode_hdr,
    _now,
    ImapReauthRequired,  # noqa: F401  -- re-exported for the caller's except sites
)

# Header set pulled per Sent item during the lightweight index pass (body is
# fetched lazily, only for an item that actually matches a vanished draft).
_HDRS = ("SUBJECT FROM TO DATE MESSAGE-ID THREAD-INDEX THREAD-TOPIC "
         "REFERENCES IN-REPLY-TO CONTENT-CLASS")

_RESP_PREFIX = re.compile(
    r"^\s*(accepted|declined|tentative|tentatively accepted|not accepted|"
    r"cancelled|canceled):\s", re.I)

# How far back to index Sent when correlating. A draft's last_seen is refreshed
# on every run, so nothing more than window_hours older than the previous run
# is ever a candidate; +10 days of slack covers a long weekend / missed runs.
_SENT_INDEX_SLACK_HOURS = 240


# --------------------------------------------------------------------------- #
#  Conversation key + body extraction
# --------------------------------------------------------------------------- #
def _conv_key(msg):
    """Stable per-conversation key for a parsed email.message.Message. Returns
    "" when nothing usable is present -- the caller then skips the item, the
    same as the COM path skipping a draft whose ConversationID is falsy."""
    ti = (msg.get("Thread-Index") or "").strip().replace(" ", "").replace("\r", "").replace("\n", "")
    if ti:
        try:
            raw = base64.b64decode(ti + "=" * (-len(ti) % 4))
            if len(raw) >= 22:
                # MS-OXOMSG: bytes[6:22] = the 16-byte conversation GUID.
                return "TIDX:" + raw[6:22].hex().upper()
        except Exception:
            pass
    refs = (msg.get("References") or "").split()
    root = refs[0] if refs else (msg.get("In-Reply-To") or "")
    root = root.strip().strip("<>").lower()
    if root:
        return "REF:" + root
    topic = _decode_hdr(msg.get("Thread-Topic") or msg.get("Subject") or "")
    while True:
        stripped = re.sub(r"(?i)^\s*(re|fw|fwd)\s*:\s*", "", topic)
        if stripped == topic:
            break
        topic = stripped
    topic = re.sub(r"\s+", " ", topic).strip().lower()
    return ("TOPIC:" + topic) if topic else ""


def _full_body(raw_bytes):
    """Full plain-text body with newlines preserved -- the IMAP equivalent of
    COM's msg.Body. NOT truncated: it feeds the redaction scan and the AI
    classifier. Prefers text/plain; falls back to a tag-stripped text/html."""
    if not raw_bytes:
        return ""
    try:
        m = email.message_from_bytes(raw_bytes)
    except Exception:
        return ""

    def _dec(part):
        try:
            raw = part.get_payload(decode=True) or b""
            return raw.decode(part.get_content_charset() or "utf-8", errors="replace")
        except Exception:
            return ""

    plain, html_txt = "", ""
    if m.is_multipart():
        for part in m.walk():
            if part.is_multipart() or part.get_filename():
                continue
            ct = part.get_content_type()
            if ct == "text/plain" and not plain:
                plain = _dec(part)
            elif ct == "text/html" and not html_txt:
                html_txt = _dec(part)
    else:
        if m.get_content_type() == "text/html":
            html_txt = _dec(m)
        else:
            plain = _dec(m)

    body = plain
    if not body.strip() and html_txt:
        body = re.sub(r"(?is)<(script|style).*?</\1>", " ", html_txt)
        body = re.sub(r"(?is)<br\s*/?>", "\n", body)
        body = re.sub(r"(?is)</p\s*>", "\n\n", body)
        body = re.sub(r"(?s)<[^>]+>", "", body)
        body = re.sub(r"&nbsp;?", " ", body)
        body = body.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    body = body.replace("\r\n", "\n").replace("\r", "\n")
    body = re.sub(r"\n{3,}", "\n\n", body).strip()
    return body


def _is_meeting_item(msg, subject, full=None):
    """Approximate COM's `Class == olMail` filter over RFC822, matching the
    exact heuristics imap_mail.py already uses for the same job."""
    if "calendarmessage" in (msg.get("Content-Class") or "").lower():
        return True
    if _RESP_PREFIX.match(subject or ""):
        return True
    if full is not None:
        try:
            if "calendar" in (full.get_content_type() or ""):
                return True
            if full.is_multipart() and any(
                    p.get_content_type() == "text/calendar" for p in full.walk()):
                return True
        except Exception:
            pass
    return False


# --------------------------------------------------------------------------- #
#  IMAP plumbing
# --------------------------------------------------------------------------- #
def _select_special(M, use_flag, name_guesses):
    """SELECT (readonly) the mailbox carrying `use_flag` (e.g. '\\Drafts'),
    falling back to well-known names. Returns the server name string, or None."""
    try:
        typ, data = M.list()
    except Exception:
        typ, data = "NO", None
    if typ == "OK" and data:
        for row in data:
            s = (row or b"").decode("utf-8", errors="replace")
            if use_flag in s:
                m = re.search(r'"([^"]*)"\s*$', s.strip()) or re.search(r'(\S+)\s*$', s.strip())
                if not m:
                    continue
                cand = m.group(1)
                try:
                    t, _ = M.select(_imap_quote(cand), readonly=True)
                    if t == "OK":
                        return cand
                except Exception:
                    pass
    for cand in name_guesses:
        try:
            t, _ = M.select(_imap_quote(cand), readonly=True)
            if t == "OK":
                return cand
        except Exception:
            continue
    return None


def _fetch_headers(M, uid):
    """-> (email.message.Message from the header subset, naive-local datetime)."""
    typ, data = M.uid("FETCH", uid,
                      f"(INTERNALDATE BODY.PEEK[HEADER.FIELDS ({_HDRS})])")
    if typ != "OK" or not data:
        return None, None
    blob, internaldate = b"", None
    for part in data:
        meta = part[0] if isinstance(part, tuple) else part
        if isinstance(part, tuple) and part[1] and b"HEADER" in (part[0] or b""):
            blob = part[1]
        im = re.search(rb'INTERNALDATE "([^"]+)"', meta or b"")
        if im:
            internaldate = im.group(1).decode("ascii", "replace")
    if not blob:
        return None, None
    msg = email.message_from_bytes(blob)
    idt = _parse_internaldate(internaldate)
    if idt is None:
        try:
            idt = email.utils.parsedate_to_datetime(msg.get("Date"))
        except Exception:
            idt = None
    if idt is not None and idt.tzinfo is not None:
        idt = idt.astimezone().replace(tzinfo=None)  # naive local -- matches COM
    return msg, idt


def _fetch_full(M, uid):
    typ, data = M.uid("FETCH", uid, "(BODY.PEEK[])")
    if typ != "OK" or not data:
        return None
    for part in data:
        if isinstance(part, tuple) and part[1]:
            return part[1]
    return None


# --------------------------------------------------------------------------- #
#  Public: Drafts snapshot  (COM parity: snapshot_drafts())
# --------------------------------------------------------------------------- #
def snapshot_drafts_imap(log=print):
    """Returns {conv_key: {conversation_topic, subject, to, body, last_seen}}
    for every current mail item in the server-side Drafts folder. Same shape
    and same semantics as draft_final_diff_capture.snapshot_drafts()."""
    token, upn = acquire_token_silent(log=log)
    M = _imap_connect(token, upn, log=log)
    snap = {}
    now_iso = datetime.now().isoformat()
    try:
        box = _select_special(M, "\\Drafts", ("Drafts", "INBOX/Drafts", "INBOX.Drafts"))
        if not box:
            log(f"draft_diff_imap - no Drafts mailbox found - snapshot empty ({_now()})")
            return snap
        typ, data = M.uid("SEARCH", None, "ALL")
        uids = data[0].split() if (typ == "OK" and data and data[0]) else []
        log(f"draft_diff_imap - Drafts mailbox {box!r}: {len(uids)} item(s) ({_now()})")
        kept = 0
        for uid in uids:
            try:
                raw = _fetch_full(M, uid)
                if not raw:
                    continue
                msg = email.message_from_bytes(raw)
                subject = _decode_hdr(msg.get("Subject"))
                if _is_meeting_item(msg, subject, msg):
                    continue
                key = _conv_key(msg)
                if not key:
                    continue
                topic = _decode_hdr(msg.get("Thread-Topic")) or subject
                snap[key] = {
                    "conversation_topic": topic,
                    "subject": subject,
                    "to": _decode_hdr(msg.get("To")),
                    "body": _full_body(raw),
                    "last_seen": now_iso,
                }
                kept += 1
            except Exception:
                continue
        scheme = {}
        for k in snap:
            scheme[k.split(":", 1)[0]] = scheme.get(k.split(":", 1)[0], 0) + 1
        log(f"draft_diff_imap - snapshot: {kept} mail draft(s) tracked; key scheme {scheme} ({_now()})")
        return snap
    finally:
        try:
            M.logout()
        except Exception:
            pass


# --------------------------------------------------------------------------- #
#  Public: Sent correlation index  (COM parity: find_sent_match())
# --------------------------------------------------------------------------- #
class SentIndex:
    """One read of \\Sent, indexed by _conv_key, for correlating vanished
    drafts. Header-only up front; the full body of a matched final is fetched
    lazily by find(). Call close() when done."""

    def __init__(self, window_hours, log=print):
        self.window_hours = window_hours
        self._log = log
        self._by_key = {}
        token, upn = acquire_token_silent(log=log)
        self._M = _imap_connect(token, upn, log=log)
        box = _select_special(self._M, "\\Sent", ("Sent Items", "Sent", "INBOX/Sent Items"))
        if not box:
            log(f"draft_diff_imap - no Sent mailbox found - correlation will find nothing ({_now()})")
            return
        floor = datetime.now() - timedelta(hours=window_hours + _SENT_INDEX_SLACK_HOURS)
        since = floor.strftime("%d-%b-%Y")
        typ, data = self._M.uid("SEARCH", None, "SINCE", since)
        uids = data[0].split() if (typ == "OK" and data and data[0]) else []
        log(f"draft_diff_imap - Sent mailbox {box!r}: indexing {len(uids)} item(s) since {since} ({_now()})")
        indexed = 0
        for uid in uids:
            try:
                msg, idt = _fetch_headers(self._M, uid)
                if msg is None or idt is None:
                    continue
                subject = _decode_hdr(msg.get("Subject"))
                if _is_meeting_item(msg, subject):
                    continue
                key = _conv_key(msg)
                if not key:
                    continue
                self._by_key.setdefault(key, []).append({
                    "uid": uid,
                    "sent_dt": idt,
                    "sent": str(idt),
                    "subject": subject,
                    "to": _decode_hdr(msg.get("To")),
                    "message_id": (msg.get("Message-ID") or "").strip(),
                })
                indexed += 1
            except Exception:
                continue
        log(f"draft_diff_imap - Sent index: {indexed} mail item(s), {len(self._by_key)} distinct conversation key(s) ({_now()})")

    def find(self, conv_key, after_dt):
        """Mirror find_sent_match(): the earliest Sent item sharing conv_key
        whose send time is in [after_dt, after_dt + window_hours]. Returns a
        dict {subject, body, to, entry_id, message_id, sent} or None. No looser
        fallback -- ambiguous/absent correlation drops the pair, never guesses."""
        rows = self._by_key.get(conv_key) or []
        end = after_dt + timedelta(hours=self.window_hours)
        cands = [r for r in rows if after_dt <= r["sent_dt"] <= end]
        if not cands:
            return None
        cands.sort(key=lambda r: r["sent_dt"])
        r = cands[0]
        body = ""
        try:
            raw = _fetch_full(self._M, r["uid"])
            body = _full_body(raw) if raw else ""
        except Exception:
            body = ""
        return {
            "subject": r["subject"],
            "body": body,
            "to": r["to"],
            "entry_id": "",           # no IMAP equivalent -- see module docstring
            "message_id": r["message_id"],
            "sent": r["sent"],
        }

    def close(self):
        try:
            if getattr(self, "_M", None):
                self._M.logout()
        except Exception:
            pass
