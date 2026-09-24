"""Per-domain last-good data handling for the Lane B briefing pipeline.

This module deliberately contains no connector calls and no dashboard logic.  It
keeps the four connector domains independent: a failed domain is either copied
as a whole from a recent last-good snapshot or is cleared and reported as
unavailable.  A domain is never assembled from a mixture of fresh and stale
items.
"""

from copy import deepcopy
from datetime import datetime, timedelta, timezone
import json
import os
import re
import shutil


MAX_CARRY_DAYS = 7

DOMAIN_FIELDS = {
    "calendar": ("calToday", "calTomorrow", "calDay2", "calDay3", "calFull", "absences"),
    "teams": ("teams",),
    "mail_inbox": ("urgent", "needs", "fyi", "low", "fyiRawCount", "suppressedCount", "mail_truncation_risk"),
    # Sent mail is an input to the AI context rather than a dashboard section.
    # It is therefore normally stored only in the local cache.
    "mail_sent": ("sent",),
}


def status_name(value):
    """Return a connector status string from either the old or new schema."""
    if isinstance(value, dict):
        return value.get("status") or "unavailable"
    return value or "unavailable"


def parse_as_of(value, *, now=None):
    """Parse ISO timestamps and the dashboard's historic human timestamp."""
    if not value:
        return None
    if isinstance(value, datetime):
        parsed = value
    else:
        text = str(value).strip().replace("Â·", "·")
        try:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError:
            parsed = None
            for fmt in ("%A %d %B · %H:%M", "%A %d %B %Y · %H:%M", "%Y%m%dT%H%M%SZ"):
                try:
                    parsed = datetime.strptime(text, fmt)
                    if "%Y" not in fmt:
                        parsed = parsed.replace(year=(now or datetime.now()).year)
                    break
                except ValueError:
                    pass
            if parsed is None:
                # A few historic files used an ASCII separator.
                match = re.match(r"^([A-Za-z]+ \d{1,2} [A-Za-z]+)(?: \d{4})?\s*[-·]\s*(\d{1,2}:\d{2})$", text)
                if match:
                    try:
                        parsed = datetime.strptime(
                            f"{match.group(1)} {match.group(2)}", "%A %d %B %H:%M"
                        )
                    except ValueError:
                        parsed = None
    if parsed is None:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=(now or datetime.now().astimezone()).tzinfo)
    return parsed.astimezone(timezone.utc)


def iso_timestamp(value, *, now=None):
    parsed = parse_as_of(value, now=now)
    return parsed.isoformat(timespec="seconds") if parsed else None


def _fields_from_doc(doc, domain):
    if not isinstance(doc, dict):
        return None
    fields = DOMAIN_FIELDS[domain]
    if domain == "mail_sent" and "sent" not in doc:
        return None
    if domain != "mail_sent" and not any(field in doc for field in fields):
        return None
    data = {}
    for field in fields:
        if field in doc:
            data[field] = deepcopy(doc[field])
    return data


def _data_has_items(data, domain):
    if not isinstance(data, dict):
        return False
    if domain == "calendar":
        return any(data.get(key) for key in ("calToday", "calTomorrow", "calDay2", "calDay3")) or any(
            day.get("items") for day in (data.get("calFull") or []) if isinstance(day, dict)
        )
    if domain == "mail_inbox":
        return any(data.get(key) for key in ("urgent", "needs", "fyi", "low"))
    if domain in ("teams", "mail_sent"):
        return bool(data.get(domain if domain == "teams" else "sent"))
    return False


def _entry_from_doc(doc, domain, *, source_override=None, require_items=True):
    data = _fields_from_doc(doc, domain)
    if source_override is not None:
        data = deepcopy(source_override)
    if data is None:
        return None
    status_doc = (doc.get("connector_status") or {}).get(domain) if isinstance(doc, dict) else None
    as_of = None
    served_by = None
    if isinstance(status_doc, dict):
        as_of = status_doc.get("as_of")
        served_by = status_doc.get("served_by")
    if not as_of and isinstance(doc, dict):
        as_of = doc.get("refreshed_at")
    if require_items and not _data_has_items(data, domain) and status_name(status_doc) != "ok":
        return None
    return {"as_of": iso_timestamp(as_of), "served_by": served_by, "data": data}


def _best_candidate(previous_doc, cache, domain):
    candidates = []
    cache_entry = ((cache or {}).get("domains") or {}).get(domain)
    if isinstance(cache_entry, dict) and cache_entry.get("data") is not None:
        candidates.append(cache_entry)
    previous_entry = _entry_from_doc(previous_doc or {}, domain)
    if previous_entry:
        candidates.append(previous_entry)
    if not candidates:
        return None
    candidates.sort(key=lambda entry: parse_as_of(entry.get("as_of")) or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
    return deepcopy(candidates[0])


def last_good_data(previous_doc, cache, domain, *, now=None, max_age_days=MAX_CARRY_DAYS):
    """Return an eligible cached/published domain entry for pipeline inputs."""
    now = now or datetime.now(timezone.utc)
    candidate = _best_candidate(previous_doc, cache, domain)
    if not candidate:
        return None
    as_of = parse_as_of(candidate.get("as_of"), now=now)
    if as_of is None or (now - as_of).total_seconds() > max_age_days * 86400:
        return None
    return candidate


def _empty_domain(briefing, domain):
    for field in DOMAIN_FIELDS[domain]:
        if field in briefing:
            old = briefing[field]
            briefing[field] = 0 if isinstance(old, (int, float)) and not isinstance(old, bool) else False if isinstance(old, bool) else []
    if domain == "mail_sent":
        briefing.pop("sent", None)


def _next_workday(day):
    day = day + timedelta(days=1)
    while day.weekday() >= 5:
        day += timedelta(days=1)
    return day


def _reproject_calendar(data, now):
    """Use calFull dates to put an older snapshot into today's four columns."""
    full = data.get("calFull") or []
    by_date = {
        str(day.get("date")): deepcopy(day.get("items") or [])
        for day in full if isinstance(day, dict) and day.get("date")
    }
    if not by_date:
        return
    dates = [now.date()]
    for _ in range(3):
        dates.append(_next_workday(dates[-1]))
    for field, day in zip(("calToday", "calTomorrow", "calDay2", "calDay3"), dates):
        data[field] = by_date.get(day.isoformat(), [])


def _status_record(status, as_of, *, served_by=None, reason=None, now=None):
    record = {"status": status}
    record["served_by"] = served_by
    normalized = iso_timestamp(as_of, now=now)
    if normalized:
        record["as_of"] = normalized
    if reason:
        record["reason"] = reason
    return record


def reconcile_domains(
    briefing,
    previous_briefing,
    cache,
    statuses,
    as_of_by_domain=None,
    *,
    enabled_domains,
    now=None,
    max_age_days=MAX_CARRY_DAYS,
    source_overrides=None,
):
    """Apply successful or failed domain results to a briefing.

    Returns ``(briefing, cache, carried_domains)``.  ``cache`` is a plain
    serialisable dictionary; the caller decides whether and where to persist
    it.  Successful domains replace their cache entry, failed domains use the
    newest eligible cache/published snapshot, and expired domains are emptied.
    """
    now = now or datetime.now(timezone.utc)
    cache = deepcopy(cache or {})
    cache.setdefault("version", 1)
    cache.setdefault("domains", {})
    statuses = statuses or {}
    as_of_by_domain = as_of_by_domain or {}
    source_overrides = source_overrides or {}
    carried = []

    for domain in enabled_domains:
        raw_status = statuses.get(domain)
        status = status_name(raw_status)
        if status == "ok":
            source = source_overrides.get(domain)
            entry = _entry_from_doc(briefing, domain, source_override=source, require_items=False)
            if entry is None:
                entry = {"data": deepcopy(source or {})}
            entry["as_of"] = iso_timestamp(as_of_by_domain.get(domain), now=now) or now.isoformat(timespec="seconds")
            cache["domains"][domain] = entry
            briefing.setdefault("connector_status", {})[domain] = _status_record(
                "ok", entry["as_of"], served_by=entry.get("served_by"), now=now
            )
            continue

        candidate = _best_candidate(previous_briefing, cache, domain)
        candidate_as_of = parse_as_of((candidate or {}).get("as_of"), now=now)
        age = (now - candidate_as_of).total_seconds() if candidate_as_of else None
        if candidate and age is not None and age <= max_age_days * 86400:
            data = deepcopy(candidate.get("data") or {})
            if domain == "calendar":
                _reproject_calendar(data, now)
            for field, value in data.items():
                briefing[field] = value
            briefing.setdefault("connector_status", {})[domain] = _status_record(
                "carried_forward", candidate.get("as_of"), served_by=candidate.get("served_by"), now=now
            )
            carried.append(domain)
            continue

        _empty_domain(briefing, domain)
        reason = "last_good_data_older_than_7_days" if candidate_as_of else "no_last_good_data"
        briefing.setdefault("connector_status", {})[domain] = _status_record(
            "unavailable", candidate.get("as_of") if candidate else None,
            served_by=candidate.get("served_by") if candidate else None,
            reason=reason, now=now
        )

    return briefing, cache, carried


def load_cache(path):
    try:
        with open(path, "r", encoding="utf-8") as handle:
            value = json.load(handle)
        return value if isinstance(value, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_cache(path, cache, *, archive_dir=None, now=None):
    """Write and validate the local cache, backing up an existing JSON first."""
    now = now or datetime.now().astimezone()
    if os.path.exists(path) and archive_dir:
        os.makedirs(archive_dir, exist_ok=True)
        stamp = now.strftime("%Y%m%d_%H%M%S_%f")
        shutil.copy2(path, os.path.join(archive_dir, f"connector_last_good_{stamp}.json"))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(cache, handle, indent=2, ensure_ascii=False)
    with open(path, "r", encoding="utf-8") as handle:
        json.load(handle)
