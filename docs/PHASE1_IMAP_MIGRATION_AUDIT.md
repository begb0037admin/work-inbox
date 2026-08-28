# Phase 1 audit — what the Outlook-COM mail pull reads, and what breaks under IMAP

**Date:** 28 August 2026 (Drew)
**Scope:** `fetch_inbox.py` Phase 1 mail pull only (inbox / VIP sweep / subfolders / Sent). Calendar (Phases 3.7/3.8) is explicitly out of scope — it stays on COM.
**Method:** direct read of `fetch_inbox.py` at commit `9a52b07` (blob `bd02b41`), full `grep` of the repo for `.Categories`, every `msg.<Prop>` / `item.<Prop>` COM attribute access enumerated, downstream consumers traced.

---

## 1. Exactly what Phase 1 reads per message

Top-level inbox loop (`for msg in restrict_date(_inbox_folder, cutoff)`), the VIP sweep, and the subfolder sweep all build the **same dict**:

| Internal key | COM source | Notes |
|---|---|---|
| `subject` | `msg.Subject` | |
| `from` | `msg.SenderName` | display name |
| `from_email` | `msg.SenderEmailAddress` | **for internal senders COM often returns an X.500 `/O=EXCHANGELABS/…` string, not SMTP** |
| `received` | `str(msg.ReceivedTime)` | `"YYYY-MM-DD HH:MM:SS+TZ"`; downstream strips the TZ |
| `is_read` | `not msg.UnRead` | |
| `has_attachments` | `msg.Attachments.Count > 0` | counts inline images too |
| `importance` | `msg.Importance` | **0 = low, 1 = normal, 2 = high** (PR_IMPORTANCE) |
| `entry_id` | `msg.EntryID` | MAPI store-scoped hex id; used as the primary key everywhere downstream |
| `kevin_is_primary_recipient` | `msg.Recipients` → per-recipient `PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)` vs `KEVIN_EMAIL`; falls back to substring match on `msg.To`; fails **open** (True) | |
| `body_preview` (unread only) | `(msg.Body or "")[:150]` | plain-text body, sliced |
| `source_folder` (subfolder items only) | `sub.FolderPath` | |

Sent sweep (`mapi.GetDefaultFolder(5).Items`) builds: `subject` (`msg.Subject`), `to` (`msg.To`), `sent` (`str(msg.SentOn)`), `body_preview` (`(msg.Body or "")[:100]`), `entry_id` (`msg.EntryID`).

`restrict_date()` filters to a **7-day** window (`cutoff = now − 7d`) with a locale-correct `dd/mm/yyyy` Restrict() string plus a 6-hour grace check and a bounded manual-iteration fallback.

Caps: top-level `MAX_UNREAD = 50`, `MAX_READ = 30`; subfolders `SUBFOLDER_MAX_UNREAD = 40`, `SUBFOLDER_MAX_READ = 20` (separate additive budget). VIP sweep is uncapped within the 7-day window.

---

## 2. Outlook **Categories** — ZERO dependence (notable de-risk)

`grep -rn -i "\.categories" --include=*.py` across the whole repo:

- **`fetch_inbox.py`: no match.** Phase 1 never reads `msg.Categories`. It never has.
- The only hits are `tools/codex_triage/mailbox_guard.py` (`it.Categories = marker` / `it.Categories = orig`) — that belongs to the **abandoned** ChatGPT-connector route (write-gate unsolvable, rejected 27–28 Aug), not the live pipeline.

`categorise()` (line ~1604) is work-inbox's **own keyword function** over `subject` / `from_email` / `is_read` / `importance` — nothing to do with Outlook Categories.

**Conclusion: nothing breaks. No IMAP equivalent needed.** The spike doc's "audit how much Phase 1 relies on categories" resolves to *not at all*.

---

## 3. **Importance / high-flag** — real dependence, cleanly recoverable from MIME headers

Consumers of `importance`:

| Site | Logic | Impact if wrong |
|---|---|---|
| `categorise()` line ~1611 | `if imp == 2: return "urgent"` | a high-importance email silently misses the Urgent tier |
| `badge_for()` line ~1636 | reads `imp` (value read; current branch logic keys on category + age, not `imp` directly) | cosmetic at most today |
| the unread-prioritisation helper (line ~789) | `high_importance = [i for i in unread if i.get("importance", 1) == 2]` | ordering only; defaults to 1 if absent |

**IMAP equivalent** (implemented in `imap_mail._importance_from_headers`): fetch `BODY.PEEK[HEADER.FIELDS (IMPORTANCE X-PRIORITY X-MSMAIL-PRIORITY)]` and map, using the same correspondence Outlook itself uses between MIME priority and PR_IMPORTANCE:

- `Importance: high` **or** `X-MSMail-Priority: High` **or** `X-Priority: 1|2` → `2`
- `Importance: low` **or** `X-MSMail-Priority: Low` **or** `X-Priority: 4|5` → `0`
- otherwise → `1`

Parity risk: an internal Exchange sender that set importance via a mechanism that doesn't emit these headers. Low. `diff_mail_pull.py` compares `importance` **and** the derived tier field-by-field, so any drift is visible before cutover.

---

## 4. `EntryID` / `openmail://` — the biggest change, **two** distinct consumers

### 4a. The dashboard opener
`open_email.py` (registered `openmail://` protocol handler) does `mapi.GetItemFromID(entry_id).Display()`. `js/app.js` `openEmail()` builds `window.location.href = 'openmail://' + entryId + '/'`. IMAP has **no EntryID** and there is no IMAP→EntryID or Message-ID→OWA-ItemID mapping without an EWS/Graph call (both unavailable at Oxford).

**Replacement:** store the internet `Message-ID` and a constructed **OWA search deep-link** `https://outlook.office.com/mail/search?query=<url-encoded Message-ID>` as `web_link` on each card. The dashboard opens it in a new tab with the **exact** validation already shipped for command-centre `sourceType=codex-graph` (`openEmailWeb`: https-only, hostname allowlist `outlook.office.com` / `outlook.office365.com`, `URL()` parse, no path/userinfo spoof tolerance). A card carries `mail_backend: "imap"` as the opener discriminator — same pattern as the 27 Aug "Open original" drafted-replies fix (discriminator field, never a dead `openmail://`). `imap_mail._owa_search_link()` builds the link; `entry_id` is set to `""` on IMAP cards so the existing `stableKey`/`draftIdentity` fallbacks (`id:` / title) still work.

**Dashboard JS change still required (NOT in this build — needs Kevin's screenshot approval):** teach `renderItems()` / the priority-card renderer / `open_email` path to branch on `mail_backend === "imap"` → `openEmailWeb`-style handler, else the current `openEmail`. Until that ships, IMAP-sourced cards would fall through to a dead `openmail://` — which is why cutover is gated on this too.

### 4b. Phase 3.9 resolution tracking (`tracked_needs_urgent`)
Lines ~3150–3290. Snapshots every live Urgent/Needs card keyed by `entry_id`; for any tracked id that scrolled out of the pull it checks three resolution signals — the first being:

```python
item = mapi.GetItemFromID(eid)
outcome = "still_open" if item.Parent.EntryID == _inbox_folder.EntryID else "moved_out"
```

Under IMAP this COM call fails → the existing per-eid `try/except` sets `outcome = "unknown"` → **fail-open, item carried**. So the current build **degrades safely** (no crash — verified by tracing the outer `try` at line ~3101 / `except` at ~3241 and the inner per-eid guard): Phase 3.9 simply stops resolving items under `imap` and carries everything, which matches its stated fail-open philosophy but would grow `tracked_needs_urgent` unbounded over time.

**Proper IMAP replacement (implemented, not yet wired — follow-up #1):** `imap_mail.message_still_in_inbox(message_ids)` runs one `UID SEARCH HEADER Message-ID "<id>"` against INBOX per tracked id. Present → `still_open` (carry); absent → `moved_out` (resolved, drop); SEARCH error → caller treats as inconclusive → carry. This is actually *more* reliable than the COM parent check (which fails-open to "unknown" frequently). Wiring it needs Phase 3.9 to key on `message_id` instead of `entry_id`; the `_cc_done_entry_ids` / `_ticked_done_entry_ids` cross-checks already operate on the stored id string and keep working as long as the key is consistent.

The other two Phase 3.9 signals are unaffected by backend: Command-Centre `tasks.json` `entryId`+`done`, and `data/ticks.json` `eid_<id>` keys — both just need the id string to match whatever key the cards carry.

---

## 5. Everything else in Phase 1 — maps cleanly

| Phase 1 element | COM | IMAP |
|---|---|---|
| 7-day window | `restrict_date()` dd/mm/yyyy Restrict() + fallback | `UID SEARCH SINCE <dd-Mon-yyyy>` + client-side re-apply of the exact cutoff against `INTERNALDATE`/`Date` |
| unread/read caps | counters | same counters, same values |
| VIP sweep | `is_vip()` on `SenderName` / `SenderEmailAddress` | same check on parsed `From:` — **IMAP gives real SMTP here, so `VIP_EMAILS` matching improves** |
| subfolder trees | `_inbox_folder.Folders` recursion, 5 named trees | `LIST "" "*"` → match `INBOX/<tree>` + `INBOX/<tree>/*` (Exchange Online hiersep `/`) |
| `Class == 43` (olMail) filter in subfolders | `getattr(msg, "Class", None)` | not needed — IMAP folders only hold RFC822 messages; meeting-item/receipt equivalents are still MIME and parse harmlessly (or are skipped on parse) |
| Sent sweep | `GetDefaultFolder(5)` | `\Sent` special-use box via `LIST`, fallback name `"Sent Items"` |
| dedup across sweeps | `captured_ids` set of `EntryID` | `captured_ids` set of `Message-ID` |

**Known IMAP gotcha to watch in parity testing:** the subfolder tree literally named **`Bi-monthly CDR/PD working group`** contains a `/`, which is also the Exchange Online IMAP hierarchy separator. Over IMAP that folder may appear as a nested path segment (`INBOX/Bi-monthly CDR/PD working group`) or with the slash server-escaped. `imap_mail._list_mailboxes()` + the `startswith(target + "/")` match may need a tweak once we see the real `LIST` output. Flagged; not blocking.

---

## 6. Summary — what actually breaks

| Item | Breaks under IMAP? | Resolution | Status |
|---|---|---|---|
| Outlook Categories | **No** — never read | none needed | ✅ confirmed |
| `importance` | Yes (no `.Importance`) | MIME `Importance:` / `X-Priority:` / `X-MSMail-Priority:` → 0/1/2 | ✅ implemented in `imap_mail` |
| `entry_id` as card key | Yes | `message_id`; `entry_id=""`; existing key-fallbacks cover it | ✅ implemented |
| Dashboard `openmail://` opener | Yes | OWA search deep-link `web_link` + `mail_backend` discriminator; reuse command-centre `openEmailWeb` validation | ⚠️ link built; **dashboard JS branch not yet written** (needs screenshot approval) |
| Phase 3.9 `GetItemFromID` resolution | Degrades (fail-open carry), does **not** crash | `imap_mail.message_still_in_inbox()` + key on `message_id` | ⚠️ helper written, **not yet wired** (follow-up #1) |
| `from_email` value shape | Changes (X.500 → real SMTP) | net improvement; parity diff will show it | ℹ️ eyeball in parity run |
| `received` value | May differ by seconds (ReceivedTime vs INTERNALDATE) | `diff_mail_pull.py` uses a 120s tolerance | ℹ️ |
| `has_attachments` | Heuristic (BODYSTRUCTURE token vs COM count incl. inline) | parity diff flags mismatches | ℹ️ |
| Calendar (3.7/3.8) | Out of scope — **stays on COM** | classic Outlook must remain runnable → WS1 keepalive stays relevant | ✅ unchanged; under `imap` a dead Outlook now degrades calendar to empty instead of failing the whole run |
