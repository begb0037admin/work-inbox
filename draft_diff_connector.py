r"""
draft_diff_connector.py -- Microsoft Outlook Email CONNECTOR backend for
draft_final_diff_capture.py.

The MAIL_BACKEND=connector path for the ongoing draft/final diff capture,
built 16 Sep 2026 (Drew) to retire imap_mail.py/reauth_imap.py from the
Oxford laptop -- "Work Inbox Laptop Draft Diff" was the last live task on
that machine hard-depending on imap_mail.py (a direct `from imap_mail import
(...)` in draft_diff_imap.py). Uses the SAME ChatGPT M365 connector already
proven live for mail_inbox/mail_sent/calendar (lane_b_call1.py) -- reuses
its low-level codex-exec + re-contamination-guard plumbing verbatim
(run_codex_json, extract_tool_calls, guard_recontamination,
ReContaminationDetected). Deliberately does NOT thread a new domain through
lane_b_call1.py's own run_domain()/EXPECTED_TOOL/_events_from_results
dispatch tables -- that machinery is tightly coupled to the already-proven
calendar/mail_inbox/mail_sent/teams paths, and this repo's own standing rule
is small isolated changes over stacked same-night edits to a live production
file. Every call here reuses domain="mail_sent" purely for
guard_recontamination()'s EXPECTED_TOOL lookup (both Sent and Drafts calls
terminate on `list_messages`, same as the real mail_sent domain) -- zero
lines of lane_b_call1.py were changed to add this capability, and
guard_recontamination()'s own allowlist is untouched.

  MAIL_BACKEND=com (default)   -> draft_final_diff_capture.py stays on Outlook
                                  COM, unchanged, this module never loads.
  MAIL_BACKEND=imap            -> draft_diff_imap.py (unchanged, still present
                                  as a rollback on any machine that keeps it).
  MAIL_BACKEND=connector       -> this module. Never imports win32com or
                                  imap_mail.

WHAT THIS REPLACES (vs draft_diff_imap.py)
  COM/IMAP                          -> connector equivalent here
  --------------------------------------------------------------------------
  Drafts folder                     -> search_messages with the connector's
                                       supported `in:drafts` mailbox filter.
  Sent Items folder                 -> search_messages with `in:sentitems`.
                                       The current connector surface exposes
                                       search_messages, not list_messages;
                                       pagination remains bounded and the
                                       safety guard remains read-only.
  msg.ConversationID / Thread-Index -> NO EQUIVALENT AVAILABLE. Confirmed
                                       live 16 Sep 2026: this connector's
                                       search_messages/fetch_message tools
                                       return a fixed field set (subject,
                                       body, bodyPreview, toRecipients,
                                       ccRecipients, sender, receivedDateTime,
                                       web_link, id, ...) with NO
                                       conversationId, internetMessageId, or
                                       raw header access, even when the model
                                       itself repeatedly tried `select:
                                       [internetMessageId, ...]` across many
                                       calls in a live trace -- the field
                                       never came back, and the eventual
                                       clean call dropped `select` entirely.
                                       Correlation here therefore uses ONLY
                                       the same subject/topic-normalisation
                                       fallback draft_diff_imap.py's own
                                       _conv_key() already falls back to as
                                       its WEAKEST tier when Thread-Index/
                                       References are absent -- this is not a
                                       new, looser tier, it is IMAP's own
                                       pre-existing worst-case tier, used
                                       here as the only tier. To offset that
                                       weaker signal, SentIndexConnector.find()
                                       additionally requires at least one
                                       shared recipient email address between
                                       the draft and the candidate sent item
                                       before accepting a match -- a real
                                       disambiguation layer neither the COM
                                       nor the IMAP path's own topic-only
                                       fallback has. Still: no looser
                                       fallback than that -- ambiguous/absent
                                       correlation drops the pair, never
                                       guesses, same discipline as both other
                                       backends.
  msg.Body                          -> body.content field (already the FULL
                                       body, HTML or text) from search_messages
                                       directly -- confirmed live 25 Sep the
                                       connector's search_messages response
                                       already includes the full body, not a
                                       preview (same finding behind the
                                       mail_sent fix in lane_b_call1.py: the
                                       extra fetch_message/fetch_messages_batch
                                       calls the model used to make were
                                       genuinely redundant).
  msg.EntryID                       -> the connector's own `id` field,
                                       provenance only, same treatment as
                                       IMAP's message_id (never used for a
                                       lookup, only recorded).

Every entry point logs a timestamp via the injected `log` callable, same
convention as draft_diff_imap.py.
"""

import os
import re
from datetime import datetime, timedelta, timezone

import lane_b_call1 as _lb

# The connector commonly returns a continuation link around 200 rows even
# when a larger top value is requested. Collect a bounded number of pages and
# refuse the pull if the connector still advertises another page after the
# bound; never rewrite the ledger from a partial folder snapshot.
DRAFTS_MAX = int(os.environ.get("WI_DRAFTDIFF_DRAFTS_MAX", "2000"))
SENT_MAX = int(os.environ.get("WI_DRAFTDIFF_SENT_MAX", "2000"))
PAGE_SIZE = max(1, int(os.environ.get("WI_DRAFTDIFF_PAGE_SIZE", "200")))
MAX_PAGES = max(1, int(os.environ.get("WI_DRAFTDIFF_MAX_PAGES", "12")))

# ---------------------------------------------------------------------------
#  Conversation key (subject/topic only -- see module docstring)
# ---------------------------------------------------------------------------
_TOPIC_STRIP_RE = re.compile(r"(?i)^\s*(re|fw|fwd)\s*:\s*")
_RESP_PREFIX = re.compile(
    r"^\s*(accepted|declined|tentative|tentatively accepted|not accepted|"
    r"cancelled|canceled):\s",
    re.IGNORECASE,
)


class ConnectorResultIncomplete(RuntimeError):
    """The connector returned a result that cannot safely represent a full folder pull."""


class ConnectorResultUnavailable(RuntimeError):
    """The connector did not produce a usable folder result; do not rewrite the ledger."""


def _normalise_topic(subject: str) -> str:
    topic = subject or ""
    while True:
        stripped = _TOPIC_STRIP_RE.sub("", topic)
        if stripped == topic:
            break
        topic = stripped
    return re.sub(r"\s+", " ", topic).strip().lower()


def _conv_key(subject: str) -> str:
    topic = _normalise_topic(subject)
    return f"TOPIC:{topic}" if topic else ""


def _recipient_emails(field) -> set:
    """field is the connector's toRecipients/to_recipients list of
    {"emailAddress"/"email_address": {"address": ...}} dicts, or occasionally
    a plain string -- tolerate both shapes defensively, same posture as
    lane_b_call1.mail_messages_to_raw()'s own field mapping."""
    out = set()
    if isinstance(field, dict):
        field = [field]
    if isinstance(field, list):
        for r in field:
            if not isinstance(r, dict):
                continue
            ea = r.get("email_address") or r.get("emailAddress") or r
            addr = (ea.get("address") or ea.get("email") or "") if isinstance(ea, dict) else ""
            addr = str(addr).strip().lower()
            if "@" in addr:
                out.add(addr)
    elif isinstance(field, str):
        # Tolerate bare addresses and Outlook's usual "Name <address>"
        # representation, but do not treat display names as email matches.
        for part in re.split(r"[;,]", field):
            match = re.search(r"<([^<>]+)>", part)
            addr = (match.group(1) if match else part).strip().lower()
            if "@" in addr:
                out.add(addr)
    return out


def _to_recipients(message: dict):
    return (message.get("to_recipients") or message.get("toRecipients")
            or message.get("to") or [])


def _is_meeting_item(message: dict, subject: str) -> bool:
    """Mirror the IMAP backend's mail-only filter for calendar messages."""
    content_class = str(
        message.get("content_class") or message.get("contentClass")
        or message.get("Content-Class") or ""
    ).lower()
    if "calendarmessage" in content_class:
        return True
    if _RESP_PREFIX.match(subject or ""):
        return True
    content_type = str(
        message.get("content_type") or message.get("contentType") or ""
    ).lower()
    if "calendar" in content_type:
        return True
    return False


def _body_text(body_field, fallback_preview: str = "") -> str:
    """The connector's `body` field is {"contentType"/"content_type":
    "html"|"text", "content": "..."}. Returns plain-ish text -- HTML tags
    stripped the same way draft_diff_imap.py's own _full_body() does for its
    html fallback, so downstream redaction/AI enrichment sees a comparable
    shape either backend. Falls back to bodyPreview/body_preview if `body`
    itself is absent (defensive only -- the live trace showed `body` always
    present)."""
    if not isinstance(body_field, dict):
        return (fallback_preview or str(body_field or "")).strip()
    content = body_field.get("content") or ""
    ctype = (body_field.get("content_type") or body_field.get("contentType") or "").lower()
    if ctype != "html":
        text = content
    else:
        text = re.sub(r"(?is)<(script|style).*?</\1>", " ", content)
        text = re.sub(r"(?is)<br\s*/?>", "\n", text)
        text = re.sub(r"(?is)</p\s*>", "\n\n", text)
        text = re.sub(r"(?s)<[^>]+>", "", text)
        text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text or (fallback_preview or "").strip()


# ---------------------------------------------------------------------------
#  Connector plumbing -- bounded mailbox-filtered search_messages pulls,
#  guard-checked
# ---------------------------------------------------------------------------
def _build_folder_prompt(display_name: str, extra_filter: str | None, top: int) -> str:
    # The current Outlook connector exposes search_messages rather than the
    # older list_mail_folders/list_messages pair. These filters were verified
    # live against the Oxford mailbox: in:drafts and in:sentitems.
    mailbox_filter = "in:drafts" if display_name.casefold() == "drafts" else "in:sentitems"
    query = f"{mailbox_filter} {extra_filter}" if extra_filter else mailbox_filter
    return (
        "Using the Microsoft Outlook Email app connector, in READ-ONLY mode, retrieve "
        f"messages from my \"{display_name}\" folder. Use only the paginated "
        "search_messages calls described below; do not add any other tool call:\n"
        f"1) Call search_messages with query=\"{query}\", size={top}, "
        "from_index=0. If has_more is true, call search_messages again with the "
        "same query and size, using the returned next_from_index. Continue for "
        f"at most {MAX_PAGES} search_messages calls and stop when has_more is false. "
        "If has_more remains after the page bound, return the last next_from_index "
        "in the raw result so the caller can reject the incomplete pull.\n"
        "Do NOT call list_mail_folders, list_messages, fetch_message, or "
        "fetch_messages_batch -- search_messages already returns the full message "
        "content needed, including the "
        "body. "
        "For each message return: subject, to recipients, received/sent date-time, "
        "the message id, the web link, and the full body. "
        "Return ONLY the raw connector result(s) as JSON (the message objects and "
        "continuation metadata), with no summary, no interpretation, and no prose. "
        "Do not use any other app or tool. Do not send, reply to, forward, move, "
        "delete, mark as read, categorise, flag, or otherwise modify any message. "
        f"{_lb.SAFETY_RULE}"
    )


def _list_folder_messages(display_name: str, *, extra_filter: str | None, top: int,
                           max_total: int, tag: str, retries: int, log=print) -> list[dict]:
    """Bounded mailbox-filtered search_messages pagination, guard-checked, with the same
    retry/backoff shape as lane_b_call1.fetch_mail_domain() (personal-account
    -only -- mail has no Edu connector, same precedent that function already
    established). Fails SOFT on anything except a genuine guard HALT: a
    missing/unavailable Drafts or Sent domain raises before the caller can
    rewrite its ledger. An incomplete or capped folder result does the same;
    an empty/partial pull must never make live drafts look vanished.
    A HALT
    (ReContaminationDetected) is re-raised, uncaught -- non-retryable by
    design, same as every other Lane B domain."""
    prompt = _build_folder_prompt(display_name, extra_filter, top)
    last_exc: Exception | None = None
    unavailable_detail = "connector did not return a usable search_messages result"
    identities = _lb.available_identity_ring()
    if not identities:
        raise ConnectorResultUnavailable(f"{display_name}: no authenticated connector identity")
    # Match lane_b_call1.fetch_domain() for Oxford mail: the checked-in ring is
    # still edu -> personal-uk -> personal-com, but personal-com is the known
    # working Outlook connector identity and therefore gets the first attempt.
    # Keep the remaining entries in their configured order for failover.
    identities = sorted(
        identities,
        key=lambda identity: 0 if identity.get("label") == "personal-com" else 1,
    )
    _lb._identity_log(
        "[identity ring] preferred personal-com for draft mail: "
        + " -> ".join(i["label"] for i in identities)
    )
    for n in range(1, retries + 1):
        identity = identities[(n - 1) % len(identities)]
        identity_label = str(identity.get("label") or "identity")
        call_tag = f"{tag}{n}-{identity_label}"
        try:
            events, raw = _lb.run_codex_json(
                _lb._prompt_for_identity(prompt, identity),
                timeout_s=_lb.CALL1_TIMEOUT_S, tag=call_tag,
                codex_home=identity["CODEX_HOME"], max_attempts=1, workload_class="high")
        except _lb.ReContaminationDetected:
            raise  # non-retryable -- propagate to the caller uncaught, same as every Lane B domain
        except Exception as e:  # noqa: BLE001 -- codex_failed / RuntimeError / transient -- retry
            last_exc = e
            log(f"draft_diff_connector - [{call_tag}] attempt failed ({e}) -- "
                f"{'trying the next identity' if n < retries else 'giving up'}")
            continue
        mismatch = _lb._account_mismatch_reason(events, identity.get("m365_account"))
        if mismatch:
            unavailable_detail = mismatch
            log(f"draft_diff_connector - [{call_tag}] {mismatch} -- trying the next identity")
            continue
        tool_calls = _lb.extract_tool_calls(events)
        status, detail = _lb.guard_recontamination(tool_calls, "mail_search")
        if status == "halt":
            raise _lb.ReContaminationDetected(f"[{call_tag}] {detail['unexpected']}")
        if status != "ok":
            unavailable_detail = str(detail)
            log(f"draft_diff_connector - [{call_tag}] connector did not return search_messages "
                f"-- trying the next identity")
            continue
        message_calls = [tc for tc in tool_calls
                         if tc["tool"].split(".")[-1] == "search_messages"]
        if not message_calls:
            unavailable_detail = f"{identity_label} returned no search_messages call"
            log(f"draft_diff_connector - [{call_tag}] {unavailable_detail} -- trying the next identity")
            continue
        if len(message_calls) > MAX_PAGES:
            raise ConnectorResultIncomplete(
                f'{display_name}: connector used {len(message_calls)} pages; bound is {MAX_PAGES}'
            )

        collected = []
        for page_index, call in enumerate(message_calls):
            res = call.get("result")
            if isinstance(res, dict):
                next_link = (res.get("next_from_index") if res.get("has_more")
                             else None)
                if next_link is None and res.get("has_more"):
                    next_link = res.get("nextFromIndex")
                page_rows = None
                for key in ("results", "value", "messages", "items"):
                    if isinstance(res.get(key), list):
                        page_rows = res[key]
                        break
            elif isinstance(res, list):
                next_link = None
                page_rows = res
            else:
                next_link = None
                page_rows = None
            if page_rows is None:
                raise ConnectorResultIncomplete(
                    f'{display_name}: page {page_index + 1} returned no message array'
                )
            collected.extend(page_rows)
            if page_index < len(message_calls) - 1 and not next_link:
                # Some connector sessions omit next_from_index from a page
                # result even though the model immediately issued the next
                # search_messages call. Accept that shape only when the next
                # call explicitly advances from_index beyond this page; if it
                # does not, retain the fail-closed partial-pull guard.
                current_from = (call.get("arguments") or {}).get("from_index")
                next_args = message_calls[page_index + 1].get("arguments") or {}
                observed_from = next_args.get("from_index")
                try:
                    current_from = int(current_from)
                    observed_from = int(observed_from)
                except (TypeError, ValueError):
                    current_from = observed_from = None
                if (current_from is None or observed_from is None
                        or observed_from <= current_from):
                    raise ConnectorResultIncomplete(
                        f'{display_name}: page {page_index + 1} had no next_from_index before another page'
                    )
                next_link = observed_from
            if page_index == len(message_calls) - 1 and next_link:
                raise ConnectorResultIncomplete(
                    f'{display_name}: has_more remains after {len(message_calls)} page(s); refusing a partial pull'
                )

        if len(collected) > max_total:
            raise ConnectorResultIncomplete(
                f'{display_name}: collected {len(collected)} items; bounded total is {max_total}'
            )
        # A continuation page can repeat the boundary row. Deduplicate only
        # by the stable connector id; rows without an id are retained.
        deduped = []
        seen_ids = set()
        for row in collected:
            row_id = row.get("id") if isinstance(row, dict) else None
            if row_id and row_id in seen_ids:
                continue
            if row_id:
                seen_ids.add(row_id)
            deduped.append(row)
        log(f"draft_diff_connector - [{call_tag}] \"{display_name}\": "
            f"{len(deduped)} item(s) across {len(message_calls)} page(s)")
        return deduped
    if last_exc is not None:
        unavailable_detail = f"last error {last_exc!r}"
    log(f"draft_diff_connector - \"{display_name}\": no usable result after {retries} "
        f"attempt(s) -- refusing to rewrite the ledger ({unavailable_detail})")
    raise ConnectorResultUnavailable(f"{display_name}: {unavailable_detail}")


# ---------------------------------------------------------------------------
#  Public: Drafts snapshot  (COM/IMAP parity: snapshot_drafts() / snapshot_drafts_imap())
# ---------------------------------------------------------------------------
def snapshot_drafts_connector(log=print) -> dict:
    """Returns {conv_key: {conversation_topic, subject, to, body, last_seen}}
    for every current mail item in the Drafts folder. Same shape and
    semantics as draft_final_diff_capture.snapshot_drafts() /
    draft_diff_imap.snapshot_drafts_imap()."""
    now_iso = datetime.now().isoformat()
    messages = _list_folder_messages("Drafts", extra_filter=None, top=PAGE_SIZE,
                                      max_total=DRAFTS_MAX,
                                      tag="draftdiff_drafts#failover", retries=_lb.CALL1_RETRIES, log=log)
    snap: dict = {}
    ambiguous_keys = set()
    for m in messages:
        if not isinstance(m, dict):
            continue
        subject = str(m.get("subject") or "")
        if _is_meeting_item(m, subject):
            continue
        key = _conv_key(subject)
        if not key:
            continue
        to_field = _to_recipients(m)
        to_addrs = _recipient_emails(to_field)
        if key in ambiguous_keys:
            continue
        if key in snap:
            # A topic-only key cannot safely represent two simultaneous drafts.
            # Remove the earlier row too; keeping either one would make a later
            # vanished-draft event indistinguishable and could poison the corpus.
            snap.pop(key, None)
            ambiguous_keys.add(key)
            continue
        to_names = []
        for r in (to_field if isinstance(to_field, list) else []):
            if not isinstance(r, dict):
                continue
            ea = r.get("email_address") or r.get("emailAddress") or r
            nm = (ea.get("name") or ea.get("address") or "") if isinstance(ea, dict) else ""
            if nm:
                to_names.append(nm)
        body_preview = m.get("body_preview") or m.get("bodyPreview") or ""
        snap[key] = {
            "conversation_topic": subject,
            "subject": subject,
            "to": ", ".join(to_names) or (m.get("to") or ""),
            "to_addrs": sorted(to_addrs),   # extra field, connector-only -- used by SentIndexConnector.find()
            "body": _body_text(m.get("body"), fallback_preview=body_preview),
            "last_seen": now_iso,
        }
    log(f"draft_diff_connector - snapshot: {len(snap)} mail draft(s) tracked, "
        f"{len(ambiguous_keys)} ambiguous duplicate topic key(s) dropped "
        f"(key scheme: subject/topic only) "
        f"({datetime.now().isoformat()})")
    return snap


# ---------------------------------------------------------------------------
#  Public: Sent correlation index  (COM/IMAP parity: find_sent_match() / SentIndex)
# ---------------------------------------------------------------------------
class SentIndexConnector:
    """One connector read of Sent Items, indexed by _conv_key (subject/topic
    only), for correlating vanished drafts. Unlike draft_diff_imap.SentIndex,
    body is already present on every indexed row (the connector's
    list_messages already returns full body -- no lazy per-match fetch is
    needed or possible via this tool surface). close() is a no-op (kept for
    interface parity with SentIndex -- there is no persistent connection to
    release on the connector path)."""

    def __init__(self, window_hours: int, log=print):
        self.window_hours = window_hours
        self._log = log
        self._by_key: dict[str, list[dict]] = {}
        floor = datetime.now(timezone.utc) - timedelta(hours=window_hours + 240)  # +10gd slack, mirrors draft_diff_imap's own
        # search_messages supports date filters on `received`; Sent Items
        # results carry the sent timestamp too, and the extra 240-hour slack
        # keeps this equivalent to the previous rolling window.
        extra_filter = f"received>={floor.strftime('%Y-%m-%d')}"
        messages = _list_folder_messages("Sent Items", extra_filter=extra_filter, top=PAGE_SIZE,
                                          max_total=SENT_MAX,
                                          tag="draftdiff_sent#failover", retries=_lb.CALL1_RETRIES, log=log)
        indexed = 0
        for m in messages:
            if not isinstance(m, dict):
                continue
            subject = str(m.get("subject") or "")
            if _is_meeting_item(m, subject):
                continue
            key = _conv_key(subject)
            if not key:
                continue
            sent_raw = m.get("sent_date_time") or m.get("sentDateTime") or m.get("receivedDateTime") or ""
            sent_dt = _lb._parse_iso_utc(sent_raw)
            if sent_dt is None:
                continue
            sent_dt = sent_dt.astimezone().replace(tzinfo=None)  # naive local, matches COM/IMAP's own comparison basis
            to_field = _to_recipients(m)
            body_preview = m.get("body_preview") or m.get("bodyPreview") or ""
            self._by_key.setdefault(key, []).append({
                "sent_dt": sent_dt,
                "sent": str(sent_dt),
                "subject": subject,
                "to_addrs": _recipient_emails(to_field),
                "to": ", ".join(
                    (r.get("email_address") or r.get("emailAddress") or {}).get("name", "")
                    for r in to_field if isinstance(r, dict)
                ) or (m.get("to") or ""),
                "body": _body_text(m.get("body"), fallback_preview=body_preview),
                "message_id": str(m.get("id") or ""),
            })
            indexed += 1
        log(f"draft_diff_connector - Sent index: {indexed} mail item(s), {len(self._by_key)} "
            f"distinct conversation key(s) ({datetime.now().isoformat()})")

    def find(self, conv_key: str, after_dt: datetime, draft_to_addrs: set | None = None) -> dict | None:
        """Mirror find_sent_match()/SentIndex.find(): the earliest Sent item
        sharing conv_key whose send time is in [after_dt, after_dt +
        window_hours]. EXTRA disambiguation layer beyond IMAP's own
        topic-only fallback (see module docstring): a candidate is only
        accepted when it shares a recipient email address with the draft,
        and exactly one candidate remains after that filter. Missing,
        non-overlapping, or still-ambiguous correlation drops the pair;
        this method never guesses."""
        rows = self._by_key.get(conv_key) or []
        end = after_dt + timedelta(hours=self.window_hours)
        cands = [r for r in rows if after_dt <= r["sent_dt"] <= end]
        if isinstance(draft_to_addrs, (list, tuple, set, frozenset)):
            draft_to_addrs = {str(a).strip().lower() for a in draft_to_addrs if "@" in str(a)}
        else:
            draft_to_addrs = _recipient_emails(draft_to_addrs)
        if not draft_to_addrs:
            if cands:
                self._log(f"draft_diff_connector - conv_key={conv_key!r}: "
                           f"{len(cands)} time-window candidate(s) but the draft has "
                           f"no usable recipient email -- dropping (no safe disambiguation)")
            return None
        with_overlap = [r for r in cands if r["to_addrs"] & draft_to_addrs]
        if not with_overlap:
            if cands:
                self._log(f"draft_diff_connector - conv_key={conv_key!r}: {len(cands)} time-window "
                           f"candidate(s) but NONE share a recipient with the draft -- dropping "
                           f"(no looser fallback)")
            return None
        if len(with_overlap) > 1:
            self._log(f"draft_diff_connector - conv_key={conv_key!r}: {len(with_overlap)} "
                       f"recipient-overlapping Sent candidates remain -- dropping as ambiguous")
            return None
        cands = with_overlap
        if not cands:
            return None
        cands.sort(key=lambda r: r["sent_dt"])
        r = cands[0]
        return {
            "subject": r["subject"],
            "body": r["body"],
            "to": r["to"],
            "entry_id": "",           # no connector equivalent -- see module docstring
            "message_id": r["message_id"],
            "sent": r["sent"],
        }

    def close(self):
        pass  # no persistent connection on the connector path -- kept for interface parity
