import json, os, base64, html, re, urllib.request, urllib.error, subprocess, time
from datetime import datetime, timedelta
from difflib import SequenceMatcher

# win32com / pywintypes / anthropic are path-specific, not universal:
#   - win32com + pywintypes: any Outlook COM work (MAIL_BACKEND=com, or the
#     COM calendar pull). A COM-free host (the Oxford laptop running
#     MAIL_BACKEND=imap) never needs them; connect_to_outlook() is already
#     wrapped under MAIL_BACKEND=imap so a missing module degrades to
#     "calendar skipped", not a crash.
#   - anthropic: only AI_BACKEND=api constructs a client (see ~line 1690).
# Guarded so a COM-free / anthropic-free box can still run. On every host that
# has them installed (both current machines do) this is byte-identical to a
# plain `import`.
try:
    import win32com.client
    import pywintypes
except ImportError:
    win32com = None
    pywintypes = None
try:
    import anthropic
except ImportError:
    anthropic = None

# com_error base for `except` clauses -- real class when pywin32 is present
# (both current machines), a never-matching stand-in otherwise so the except
# sites below stay valid on a COM-free box.
_COM_ERROR = pywintypes.com_error if pywintypes is not None else OSError

# Suppress Windows git gc --auto interactive prompts
subprocess.run(["git", "config", "gc.auto", "0"], capture_output=True,
               cwd=os.path.dirname(os.path.abspath(__file__)))

# Every run must print a clear timestamp so pasted console/log output is
# self-dating -- a pasted traceback with no timestamp anywhere made an
# already-fixed incident look like a fresh failure (Kevin, 12 Aug 2026).
def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

log("fetch_inbox.py run started")

# --------------------------------------------------------------------------- #
#  Failure-toast helper. Defined here (near the top) so connect_to_outlook()
#  can use it -- moved up from further down the file on 2026-08-28 (Drew).
#  Reuses the exact same mechanism already wired to the run
#  (Show-TaskNotification.ps1 / BurntToast). Writes a dedicated one-line
#  detail file so the toast text is deterministic regardless of what the
#  shared run log looks like by the time the toast script reads it.
#  Best-effort only: a failure to raise the toast must never mask or replace
#  the exception the caller is already handling, and must never crash the run.
# --------------------------------------------------------------------------- #
NOTIFY_SCRIPT_PATH = r"D:\OneDrive - lelitte.com\Desktop\Show-TaskNotification.ps1"

def _notify_phase_failure(task_name, detail):
    try:
        detail_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "phase_failure_last.log")
        with open(detail_path, "w", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {task_name} failed: {detail}\n")
        if os.path.exists(NOTIFY_SCRIPT_PATH):
            subprocess.run(
                ["powershell.exe", "-NoProfile", "-WindowStyle", "Hidden", "-File", NOTIFY_SCRIPT_PATH,
                 "-Status", "Failure", "-TaskName", task_name, "-LogPath", detail_path],
                timeout=20, capture_output=True
            )
        else:
            print(f"WARNING: failure toast skipped for '{task_name}' - notification script not found at {NOTIFY_SCRIPT_PATH}")
    except Exception as notify_err:
        print(f"WARNING: failure toast for '{task_name}' could not be sent - {notify_err}")

def _imap_reauth_toast_due(min_interval_s=3600):
    """1-per-hour gate for the IMAP mail re-auth toast, mirroring the WS1
    Classic-Outlook-keepalive stamp mechanism (a .stamp file under
    %LOCALAPPDATA%\\WorkInboxAI, 1/hour). Returns True if a toast should be
    raised now, and touches the stamp when it does. Fails OPEN (True) on any
    stamp error -- never suppress a real re-auth alert over bookkeeping."""
    try:
        d = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "WorkInboxAI")
        os.makedirs(d, exist_ok=True)
        stamp = os.path.join(d, "imap_reauth_toast.stamp")
        if os.path.exists(stamp) and (time.time() - os.path.getmtime(stamp)) < min_interval_s:
            return False
        with open(stamp, "w", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        return True
    except Exception:
        return True

GITHUB_REPO = "begb0037admin/work-inbox"
GITHUB_PATH = "data/briefing.json"
GITHUB_PAT  = os.environ.get("GITHUB_PAT", "")
GITHUB_TIMEOUT = 30

# --------------------------------------------------------------------------- #
#  AI backend selection + parallel-validation mode  (added 2026-08-27, Drew)
#  Codex-connector migration pivoted to headless Claude Code -- see
#  docs/CLAUDE_CODE_HEADLESS_SCOPE.md and docs/CLAUDE_CODE_BACKEND.md.
#
#  AI_BACKEND=api          (default) -- unchanged: metered Anthropic API,
#                          exact behaviour as before this change.
#  AI_BACKEND=claude_code  -- the 5 claude-haiku-4-5 calls go through
#                          `claude -p` (Claude Code, subscription auth), tools
#                          disabled, no MCP. Same model, same verbatim prompts.
#
#  WI_AI_PARALLEL=1        -- parallel-validation run: do ALL the COM + AI work
#                          but write claude_*-prefixed LOCAL files only, push
#                          NOTHING, mutate no shared ledger / no Command Centre
#                          sync. For diffing the claude_code output against the
#                          live api pipeline before any cutover.
#
#  Account selection / dual-account failover (kevin@ primary, hope@ overflow):
#  WI_CLAUDE_CONFIG_DIR           -- CLAUDE_CONFIG_DIR for the primary account
#  WI_CLAUDE_CONFIG_DIR_FALLBACK  -- CLAUDE_CONFIG_DIR for the overflow account;
#                          on a usage-limit error the call is retried once on
#                          this account. Leave unset to disable failover.
# --------------------------------------------------------------------------- #
AI_BACKEND  = os.environ.get("AI_BACKEND", "api").strip().lower()
# --------------------------------------------------------------------------- #
#  Mail-pull backend selection  (added 2026-08-28, Drew)
#  MAIL_BACKEND=com   (default) -- Phase 1 mail pull (inbox / VIP sweep /
#                     subfolders / Sent) via Outlook COM, byte-identical to
#                     before this change.
#  MAIL_BACKEND=imap  -- that same mail pull via IMAP+OAuth2 (imap_mail.py).
#                     Calendar phases 3.7/3.8 stay on Outlook COM regardless.
#                     Do NOT set this on the scheduled task until
#                     diff_mail_pull.py parity is clean over several cycles
#                     AND Kevin has given a fresh explicit go-ahead. See
#                     docs/MAIL_BACKEND_MIGRATION_PLAN.md.
#  WI_MAIL_PARALLEL=1 -- dump the raw COM/IMAP mail lists to data/parallel/
#                     for diff_mail_pull.py and push / mutate NOTHING
#                     (folds into the same no-write posture as WI_AI_PARALLEL).
# --------------------------------------------------------------------------- #
MAIL_BACKEND  = os.environ.get("MAIL_BACKEND", "com").strip().lower()
MAIL_PARALLEL = os.environ.get("WI_MAIL_PARALLEL", "").strip().lower() in ("1", "true", "yes")
# --------------------------------------------------------------------------- #
#  Calendar-source backend  (added 2026-08-29, Drew -- laptop migration Phase 3)
#  CAL_BACKEND=com        (default) -- calendar phases 3.7/3.8 pull the primary
#                         + "People Department - HR Systems" calendars via
#                         Outlook COM, byte-identical to before.
#  CAL_BACKEND=connector  -- pull the calendar via the Lane B ChatGPT M365
#                         connector (codex exec). NOT IMPLEMENTED YET -- Lane B
#                         is not built until 1 Sept (see docs/
#                         LANE_B_TEAMS_CAL_DESIGN.md). Until then it means "no
#                         COM calendar source this run": calendar phases go
#                         empty + warning, and classic Outlook is NOT opened.
# --------------------------------------------------------------------------- #
_CAL_BACKEND_REQ = os.environ.get("CAL_BACKEND", "com").strip().lower()
CAL_BACKEND = _CAL_BACKEND_REQ if _CAL_BACKEND_REQ in ("com", "connector") else "com"
CAL_CONNECTOR_NYI = (CAL_BACKEND == "connector")

#  WI_BRIDGE_ALLOW_EMPTY_CALENDAR=1 -- "laptop bridge" mode. The laptop has no
#  calendar source (no classic Outlook; Lane B not built), so calendar summaries
#  AND calendar-derived absences are legitimately empty this run. When set, the
#  Phase-4 safe-write guard does NOT veto the push for "same-day calendar
#  summaries would be removed" / "calendar summaries dropped" / "same-day
#  absences would be cleared" -- it downgrades each to a WARNING. Every other
#  safe-write check (context degradation, etc.) is unchanged. Set by
#  "Run Laptop Bridge Briefing.ps1". See docs/LAPTOP_MIGRATION_PLAN.md.
BRIDGE_ALLOW_EMPTY_CALENDAR = os.environ.get(
    "WI_BRIDGE_ALLOW_EMPTY_CALENDAR", "").strip().lower() in ("1", "true", "yes")
if CAL_CONNECTOR_NYI:
    log("Calendar backend: 'connector' requested -- NOT IMPLEMENTED yet (Lane B "
        "lands 1 Sept, see docs/LANE_B_TEAMS_CAL_DESIGN.md). Calendar will be "
        "empty this run; mail briefing continues; Outlook COM will not be used.")
AI_PARALLEL = (os.environ.get("WI_AI_PARALLEL", "").strip().lower() in ("1", "true", "yes")
               or MAIL_PARALLEL)
PUSH_ENABLED = bool(GITHUB_PAT) and not AI_PARALLEL
_AI_OUT_PREFIX = "claude_" if AI_PARALLEL else ""
CLAUDE_BIN          = os.environ.get("WI_CLAUDE_BIN", "claude")
CLAUDE_CFG_PRIMARY  = os.environ.get("WI_CLAUDE_CONFIG_DIR", "").strip()
CLAUDE_CFG_FALLBACK = os.environ.get("WI_CLAUDE_CONFIG_DIR_FALLBACK", "").strip()
_AI_CALL_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ai_backend_usage.jsonl")
_ai_call_seq = 0

log(f"AI backend: {AI_BACKEND}"
    + (f"  [PARALLEL VALIDATION MODE -- local claude_* files only, no push]" if AI_PARALLEL else "")
    + (f"  primary_cfg={CLAUDE_CFG_PRIMARY or '(default)'}" if AI_BACKEND == "claude_code" else "")
    + (f"  fallback_cfg={CLAUDE_CFG_FALLBACK}" if (AI_BACKEND == "claude_code" and CLAUDE_CFG_FALLBACK) else ""))
log(f"Mail backend: {MAIL_BACKEND}"
    + ("  [WI_MAIL_PARALLEL -- raw mail dumps to data/parallel/, no push]" if MAIL_PARALLEL else ""))
log(f"Calendar backend: {CAL_BACKEND}"
    + ("  [WI_BRIDGE_ALLOW_EMPTY_CALENDAR -- empty calendar/absences will not veto the Phase 4 push]"
       if BRIDGE_ALLOW_EMPTY_CALENDAR else ""))


class _AIText:
    """Minimal stand-in for an anthropic Message: exposes .content[0].text and
    .usage so the five existing call sites need no other change."""
    __slots__ = ("content", "usage")
    def __init__(self, text, usage):
        self.content = [type("_Blk", (), {"text": text})()]
        self.usage = usage


def _looks_like_usage_limit(s):
    s = (s or "").lower()
    return any(k in s for k in (
        "usage limit", "rate limit", "rate_limit", "rate-limit", " 429", "\"429\"",
        "quota", "overloaded", "try again later", "capacity", "exceeded your",
        "usage cap", "limit reached"))


def _claude_code_once(model, system, user, timeout_s, cfg_dir):
    cmd = [
        CLAUDE_BIN, "-p",
        "--model", model,
        "--system-prompt", system,
        "--exclude-dynamic-system-prompt-sections",
        # triage is text-in / JSON-out: no tool use, no MCP, no write path.
        "--disallowedTools",
        "Bash,Edit,Write,Read,Glob,Grep,WebFetch,WebSearch,NotebookEdit,Task,TodoWrite,SlashCommand",
        "--strict-mcp-config",
        "--mcp-config", '{"mcpServers":{}}',
        "--permission-mode", "default",
        "--no-session-persistence",
        "--output-format", "json",
    ]
    # Clean environment: no metered API key (force subscription billing), and
    # drop the parent Claude Code session's IPC/session vars so this nested
    # `claude -p` starts a fully independent session.
    env = {k: v for k, v in os.environ.items()
           if k != "ANTHROPIC_API_KEY"
           and not k.startswith("CLAUDE_CODE")
           and k not in ("CLAUDECODE", "CLAUDE_PID", "CLAUDE_EFFORT", "AI_AGENT")}
    if cfg_dir:
        env["CLAUDE_CONFIG_DIR"] = cfg_dir
    p = subprocess.run(cmd, input=user, capture_output=True, text=True,
                       timeout=timeout_s, env=env)
    if p.returncode != 0:
        raise RuntimeError(f"claude -p rc={p.returncode}: {(p.stderr or p.stdout or '')[-400:]}")
    try:
        obj = json.loads(p.stdout)
    except Exception:
        raise RuntimeError(f"claude -p produced non-JSON stdout: {p.stdout[:400]}")
    if isinstance(obj, dict) and obj.get("is_error"):
        raise RuntimeError(f"claude -p is_error: {str(obj.get('result'))[:400]}")
    return obj


_CC_LAST_ACCOUNT = "primary"


def _claude_code_call(system, user, timeout_s):
    """Run ONE `claude -p` with the dual-account (kevin@ primary -> hope@
    overflow) failover loop. Returns the parsed `claude -p` JSON dict.
    Raises RuntimeError if every account/try is exhausted.
    A subprocess stall past `timeout_s` is treated as a rate-limit signal and
    switches account (seen on a Pro plan under load, 27 Aug -- a stall, not a
    crash)."""
    global _CC_LAST_ACCOUNT
    attempts = [("primary", CLAUDE_CFG_PRIMARY)]
    if CLAUDE_CFG_FALLBACK:
        attempts.append(("fallback", CLAUDE_CFG_FALLBACK))
    last_err = None
    for acct_label, cfg_dir in attempts:
        for tryno in (1, 2):
            t0 = datetime.now()
            try:
                obj = _claude_code_once("claude-haiku-4-5", system, user, timeout_s, cfg_dir)
                _CC_LAST_ACCOUNT = acct_label
                dur = (datetime.now() - t0).total_seconds()
                log(f"claude_code call account={acct_label} try={tryno} wall={dur:.1f}s ok")
                return obj
            except subprocess.TimeoutExpired as e:
                last_err = e
                log(f"claude_code call account={acct_label} try={tryno} TIMED OUT after "
                    f"{timeout_s:.0f}s -- treating as rate-limit stall, switching account.")
                break   # a stall this long => try the fallback account
            except Exception as e:  # noqa: BLE001
                last_err = e
                msg = str(e)
                log(f"claude_code call account={acct_label} try={tryno} FAILED: {msg[:220]}")
                if _looks_like_usage_limit(msg):
                    break   # switch account rather than retry the same one
                time.sleep(8)
    raise RuntimeError(f"claude_code backend exhausted (last: {last_err})")


# --------------------------------------------------------------------------- #
#  claude_code backend: ONE combined `claude -p` call for all five phases
#  (added 2026-08-27, Drew -- mitigation #2 in docs/CLAUDE_CODE_BACKEND.md,
#  design ported from tools/codex_triage/build_call2_brief.py). Fired once,
#  early (right after Phase 3 card-building, where every phase payload can be
#  assembled); each phase block below then reads its slice from _CC_COMBINED
#  via _ai_create(_phase=...). Removes 4x the per-`claude -p` cache-creation +
#  harness overhead. The `api` backend is completely unaffected -- it still
#  makes five separate client.messages.create() calls.
# --------------------------------------------------------------------------- #
_CC_COMBINED = None            # dict of the 5 phase slices, or None on failure
_CC_COMBINED_USAGE = {}        # usage dict from the single combined call
_PHASE_KEY = {
    "context":       "context_phase",
    "email_summary": "email_summary_phase",
    "task_triage":   "task_triage_phase",
    "task_summary":  "task_summary_phase",
    "calendar_prep": "calendar_prep_phase",
}

# Verbatim phase system prompts, hoisted to module scope so the early combined
# `claude -p` call can use them. The four Phase-3.x blocks below assign their
# own name FROM these (EMAIL_SUMMARY_SYSTEM = _SYS_EMAIL_SUMMARY, etc.) so
# there is exactly one source of truth and the api path is unchanged.
_SYS_EMAIL_SUMMARY = (
    "You are Kevin's inbox briefing assistant at Oxford University Personnel Services.\n"
    "For each email, write ONE concise sentence summarising what it is actually about and "
    "what, if anything, Kevin needs to do. Do not just repeat the subject line or copy the "
    "opening words verbatim - genuinely summarise the content. Be specific - use names, "
    "dates and case numbers where present. Plain ASCII punctuation only.\n"
    "Also decide needs_reply: true if this email genuinely calls for Kevin to send a reply "
    "(a question, a request, something someone is waiting to hear back on), false if it just "
    "needs him to read it, take an offline action, or do nothing at all (e.g. a system "
    "notification, an FYI, a failed-import alert, a case update that doesn't ask him anything "
    "directly).\n"
    "Also decide no_action_needed: true ONLY if Kevin genuinely has nothing to do with this "
    "email at all - a pure FYI, an automated notification, a colleague-to-colleague thread "
    "he's just cc'd on for visibility, a status update that doesn't need him to act. false if "
    "needs_reply is true, OR if Kevin needs to do anything else even without writing a reply - "
    "review something, approve something, action a request personally, follow up with someone, "
    "or respond to a meeting invite that's specifically asking for his availability/decision. "
    "no_action_needed must always be false whenever needs_reply is true - never set both true.\n"
    "Weigh two extra signals given for each email:\n"
    "- kevin_is_primary_recipient: false means Kevin was only cc'd, not directly addressed. "
    "Default toward needs_reply: false for cc-only threads UNLESS the content clearly still "
    "asks Kevin himself something directly (e.g. someone names him and asks a question even "
    "on a cc'd thread) - don't flip mechanically, use judgement. Being cc'd does NOT by itself "
    "mean no_action_needed: true - a cc'd thread can still need Kevin to review, approve, or "
    "follow up on something even without a direct question. Only set no_action_needed: true "
    "for a cc'd thread when it's genuinely visibility-only (e.g. two other people confirming "
    "something between themselves that doesn't involve a decision or action of Kevin's) - if "
    "in doubt whether a cc'd thread needs Kevin to do something, leave no_action_needed: "
    "false.\n"
    "- age_days: how many days old the email is. Default toward needs_reply: false for "
    "anything genuinely old (multiple weeks+) - an unanswered thread that old is more likely "
    "already resolved elsewhere than still genuinely awaiting Kevin's reply.\n"
    "Return ONLY a valid JSON object mapping the given short id to an object with 'summary', "
    "'needs_reply' and 'no_action_needed' - no preamble, no markdown.\n"
    'Example: {"0": {"summary": "Marie confirms funding approved for SBS exclusion from '
    'the DSE feed; no action needed from Kevin.", "needs_reply": false, "no_action_needed": '
    'true}, "1": {"summary": "James is asking whether the FA KPI meeting can move to '
    'Thursday.", "needs_reply": true, "no_action_needed": false}, "2": {"summary": '
    '"Christopher forwards the tender evaluation pack and needs Kevin to review and sign off '
    'before Friday.", "needs_reply": false, "no_action_needed": false}}'
)

_SYS_TRIAGE = (
    "You are Kevin's task triage assistant at Oxford University Personnel Services.\n"
    "You receive his existing Command Centre task list, his recent action-required received emails, and emails Kevin himself sent (direction: sent).\n"
    "Identify:\n"
    "1. new_tasks - emails that represent real, actionable work for Kevin that is NOT covered by any existing task. Max 12. "
    "Do not be over-cautious: if an email asks Kevin for something, or commits him to something, and no existing task covers it, propose it. "
    "It is better to propose a task Kevin dismisses in one click than to leave real work invisible.\n"
    "If an email concerns work that any existing task already covers - even partially, even if you would mention that task in your description - it belongs in task_updates with that task's id, NEVER in new_tasks.\n"
    "2. task_updates - emails that are progress, replies or new information on an EXISTING task. Max 20. "
    "A task_update must clearly concern that specific task - same case number, same named project, or same people AND topic. "
    "If no existing task is a clear match, do NOT force one: either propose it under new_tasks or omit it entirely.\n"
    "Return ONLY a valid JSON object - no preamble, no markdown, no code fences. Plain ASCII punctuation only.\n"
    "{\n"
    '  "new_tasks": [{"email_n": <n>, "title": "<short imperative task title>", "tier": "today|tomorrow|week", "description": "<2-3 sentences: what the work is and why, drawn from the email>"}],\n'
    '  "task_updates": [{"email_n": <n>, "task_id": "<existing task id>", "note": "<one sentence: what this email adds to the task>"}]\n'
    "}\n"
    'Rules: tier "today" only if the deadline is today or overdue; "tomorrow" if it must happen the next working day; otherwise "week". '
    "Never invent case numbers or names. Automated notifications, newsletters, calendar "
    "accept/decline messages and out-of-office replies are never tasks. "
    "Use direction=sent emails to log Kevin's own actions on existing tasks as task_updates "
    "(e.g. 'Kevin replied to Reenu with the requested staff list') so the action log shows "
    "both sides of the conversation. Never propose a new task for work that a sent email "
    "shows Kevin has already handled."
)

_SYS_TASK_SUMMARY = (
    "You are Kevin's task briefing assistant at Oxford University Personnel Services.\n"
    "For each task, write a 1-2 sentence status summary: current state, what needs to happen next, any blockers.\n"
    "Be specific - use names, dates and case numbers from the data. Plain ASCII punctuation only.\n"
    "Return ONLY a valid JSON object mapping task id to summary string - no preamble, no markdown.\n"
    "Example: {\"task-001\": \"Awaiting response from Jane Smith re HRIS migration. Next: chase by Friday 20 Jun.\"}"
)

_SYS_CAL = (
    "You are Kevin's briefing assistant at Oxford University HR Systems.\n"
    "For each meeting, write 2-3 concise sentences of prep context Kevin needs before walking in.\n"
    "Where 'prev_meeting_notes' is provided, use it as your primary source -- it is the AI summary from the last time this meeting ran.\n"
    "Prioritise: carry-forwards and open actions from last time, any live decision or blocker, who Kevin needs to speak to, and the most useful detail Kevin should remember.\n"
    "Plain ASCII punctuation only. No filler like 'This meeting is about...'. Be direct and specific.\n"
    "Return ONLY valid JSON: {\"day_idx\": \"2-3 concise sentences\"} where day_idx is 'today_0', 'today_1', 'tomorrow_0' etc.\n"
    "Example: {\"today_0\": \"Pick up the evaluation scoring from last week -- Helen still needs a decision on weightings. Confirm whether James has resolved the reporting extract and agree the next owner before Friday.\"}"
)

# Command Centre + Granola config -- hoisted to module scope so the early
# combined call can use them (originals below just call the loader).
COMMAND_CENTRE_REPO = "begb0037admin/command-centre"
COMMAND_CENTRE_PATH = "data/tasks.json"
# Auto-create Command Centre tasks from inbox suggestions without review.
# Enabled 2026-08-02 by Kevin's explicit instruction, overriding the default-off
# stance taken because command-centre/CLAUDE.md reserves new-task creation as
# his approval authority. Each promoted task still carries an
# "origin": "inbox-auto" tag and an auto-created action-log entry so promoted
# tasks stay distinguishable from ones Kevin created directly.
AUTO_PROMOTE_NEW_TASKS = True
GRANOLA_API_KEY = os.environ.get("GRANOLA_API_KEY", "")
_granola_context = {}          # "day_idx" -> {"note_title": str, "summary": str}
_all_day_candidates = []       # populated for claude_code by _cc_build_cal_candidates_early()
priorities_today, priorities_tomorrow, priorities_week = [], [], []
cc_content = []
_cc_priorities_loaded = False


def _granola_keywords(title):
    t = re.sub(r'\b\d{1,2}/\d{2}\b', '', title)   # remove DD/MM dates
    t = re.sub(r'\b\d{4}\b', '', t)                # remove years
    t = re.sub(r'[—\-&]', ' ', t)             # dashes and ampersands to spaces
    t = re.sub(r'[^\w\s]', '', t)                  # strip remaining punctuation
    return set(w.lower() for w in t.split() if len(w) >= 2)


def _granola_fetch(url):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {GRANOLA_API_KEY}"})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode())


def _non_all_day_candidates(items, day_label):
    # "idx" (sequential 0-based within its day) is the ONLY index shown to the
    # model; "real_idx" (true position in the day's item list) is used only for
    # the local write-back. Fixes the calendar-summary offset bug (4 Aug 2026).
    candidates = []
    model_idx = 0
    for real_idx, c in enumerate(items):
        if c.get("time", "").lower() == "all day":
            continue
        candidates.append({
            "idx": model_idx, "real_idx": real_idx, "day": day_label,
            "time": c["time"], "title": c["title"], "organizer": c.get("sub", "")
        })
        model_idx += 1
    return candidates


def _cc_load_priorities():
    """Load Command Centre tasks.json -> cc_content + priorities_{today,
    tomorrow,week}. Verbatim logic from the Phase 3 'Priority actions' block;
    idempotent so the early claude_code call and the original site can both
    invoke it."""
    global _cc_priorities_loaded, cc_content
    global priorities_today, priorities_tomorrow, priorities_week
    if _cc_priorities_loaded:
        return
    _cc_priorities_loaded = True
    priorities_today, priorities_tomorrow, priorities_week = [], [], []
    cc_content = []
    try:
        cc_url = f"https://api.github.com/repos/{COMMAND_CENTRE_REPO}/contents/{COMMAND_CENTRE_PATH}"
        cc_headers = {
            "Authorization": f"token {GITHUB_PAT}",
            "Content-Type":  "application/json",
            "User-Agent":    "work-inbox-script"
        }
        cc_req = urllib.request.Request(cc_url, headers=cc_headers)
        with urllib.request.urlopen(cc_req, timeout=GITHUB_TIMEOUT) as r:
            cc_data    = json.loads(r.read())
            cc_content = json.loads(base64.b64decode(cc_data["content"]).decode("utf-8"))
        task_list = cc_content if isinstance(cc_content, list) else cc_content.get("tasks", [])
        for task in task_list:
            if task.get("done"):
                continue
            tier = task.get("tier", "")
            entry = {
                "id":          task.get("id", ""),
                "text":        task.get("title", ""),
                "description": task.get("description", ""),
                "actions":     task.get("actions", []),
                "source":      task.get("source", ""),
                "dateType":    "red" if tier == "today" else "orange"
            }
            if tier == "today":
                priorities_today.append(entry)
            elif tier == "tomorrow":
                priorities_tomorrow.append(entry)
            elif tier == "week":
                priorities_week.append(entry)
        print(f"Command Centre loaded - today:{len(priorities_today)} tomorrow:{len(priorities_tomorrow)} week:{len(priorities_week)}")
    except Exception as e:
        print(f"WARNING: Could not load Command Centre tasks - {e}")
        priorities_today, priorities_tomorrow, priorities_week = [], [], []


def _cc_build_cal_candidates_early():
    """Build _all_day_candidates from the RAW per-day calendar lists (already
    leave-filtered in Phase 2) so the calendar-prep payload is ready for the
    single combined call BEFORE the absence-detection block that build_cal_items
    depends on has run. real_idx = position in the start-sorted day list = the
    index build_cal_items produces for cal_<day>_items (same sort, same length,
    no absence substitution because leave items are pre-filtered), so Phase
    3.8's write-back onto cal_<day>_items stays correct."""
    global _all_day_candidates

    def _adapt(day_list):
        # Mirror build_cal_items()'s per-item shape (time as HH:MM or "All day",
        # title from subject, sub from organizer) and its start-sort, so the
        # model sees the same meeting times the api path shows it and real_idx
        # lines up with cal_<day>_items for Phase 3.8's write-back.
        out = []
        for it in sorted(day_list, key=lambda x: x.get("start", "")):
            if it.get("all_day"):
                tstr = "All day"
            else:
                try:
                    tstr = datetime.fromisoformat(it.get("start", "")).strftime("%H:%M")
                except Exception:
                    tstr = ""
            out.append({
                "time":  tstr,
                "title": it.get("subject", ""),
                "sub":   it.get("organizer", "") or "",
            })
        return out

    _all_day_candidates = (
        _non_all_day_candidates(_adapt(cal_today), "today") +
        _non_all_day_candidates(_adapt(cal_tomorrow), "tomorrow") +
        _non_all_day_candidates(_adapt(cal_day2), "day2") +
        _non_all_day_candidates(_adapt(cal_day3), "day3")
    )


def _cc_fetch_granola(candidates):
    """Populate _granola_context from recent Granola notes for the given
    idx-fixed candidate list. Verbatim match/lookback logic from Phase 3.7b;
    shared by the early claude_code call and the original site."""
    global _granola_context
    if not GRANOLA_API_KEY:
        print("Phase 3.7 skipped - GRANOLA_API_KEY not set")
        return
    try:
        _lookback = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%dT%H:%M:%SZ")
        _g_notes  = _granola_fetch(
            f"https://public-api.granola.ai/v1/notes?created_after={_lookback}").get("notes", [])
        for cal_item in candidates:
            cal_kw = _granola_keywords(cal_item["title"])
            if not cal_kw:
                continue
            best_note, best_score = None, 0
            for note in _g_notes:
                score = len(cal_kw & _granola_keywords(note.get("title", "")))
                if score > best_score:
                    best_score, best_note = score, note
            if best_note and best_score >= 1:
                detail   = _granola_fetch(
                    f"https://public-api.granola.ai/v1/notes/{best_note['id']}?include=transcript")
                _raw_sum = detail.get("summary") or ""
                if isinstance(_raw_sum, dict):
                    summary = (_raw_sum.get("text") or _raw_sum.get("content") or "").strip()
                else:
                    summary = str(_raw_sum).strip()
                if not summary:
                    summary = (detail.get("summary_text") or detail.get("summary_markdown") or "").strip()
                if summary:
                    _granola_context[f"{cal_item['day']}_{cal_item['idx']}"] = {
                        "note_title": best_note.get("title", ""), "summary": summary[:1500]}
        print(f"Phase 3.7 done - Granola context for {len(_granola_context)} meetings")
    except Exception as e:
        print(f"WARNING: Phase 3.7 Granola fetch failed - {e}")


def _p2_finalise(context, subtitle):
    """Phase 2 tail: same-day preservation + Outlook-data fallback. Verbatim
    from the original Phase 2 block; called from the api path in place and
    from the claude_code path after the combined call."""
    if same_briefing_date(existing_briefing, today_str):
        if existing_briefing.get("context"):
            context = existing_briefing["context"]
            print("Phase 2 preservation - reused existing same-day context")
        if existing_briefing.get("subtitle"):
            subtitle = existing_briefing["subtitle"]
    if not context:
        context = build_fallback_context(inbox, cal_today, cal_tomorrow)
        print("Phase 2 fallback - generated context directly from Outlook data")
    if not subtitle:
        subtitle = build_fallback_subtitle(inbox)
    return context, subtitle


def _cc_run_combined():
    """Assemble ONE `claude -p` call covering all five judgement phases and
    parse the combined JSON into _CC_COMBINED. Payloads are built with the
    same deterministic logic each phase block rebuilds for its own downstream
    index-mapping, so the response keys line up. claude_code backend only."""
    global _CC_COMBINED, _CC_COMBINED_USAGE

    # 1. context -- reuse Phase 2's already-built SYSTEM / USER verbatim
    p1_sys, p1_user = SYSTEM, USER

    # 2. email summaries
    def _age_days(card):
        try:
            rec_dt = datetime.fromisoformat(card.get("received_raw", "").split("+")[0].split(" (")[0].strip())
            return (datetime.now() - rec_dt).days
        except Exception:
            return None
    _efs = [
        {
            "id":      str(i),
            "subject": c["subject"],
            "from":    c["from"],
            "preview": (c.get("sub") or "")[:250],
            "kevin_is_primary_recipient": c.get("kevin_is_primary_recipient", True),
            "age_days": _age_days(c),
        }
        for i, c in enumerate(summary_candidates)
    ]
    p2_user = f"Today is {today_str}.\n\nEMAILS:\n{json.dumps(_efs, indent=1, ensure_ascii=True)}"

    # 3. task triage
    _tl = cc_content if isinstance(cc_content, list) else cc_content.get("tasks", [])
    _tsum = [
        {"id": t.get("id", ""), "title": t.get("title", ""),
         "description": (t.get("description") or "")[:300], "emailRef": t.get("emailRef", "")}
        for t in _tl
    ]
    _ec = []
    for m in inbox:
        if categorise(m) in ("urgent", "needs"):
            _ec.append({
                "subject":      m.get("subject", ""),
                "from":         m.get("from", ""),
                "received":     (m.get("received", "") or "")[:16],
                "body_preview": re.sub(r"<\?\s*https?://\S+>?", "[link]", (m.get("body_preview") or ""))[:150],
                "entry_id":     m.get("entry_id", ""),
            })
    for s in sent[:30]:
        _ec.append({
            "subject":      s.get("subject", ""),
            "from":         "Kevin (sent to: " + (s.get("to") or "") + ")",
            "received":     (s.get("sent", "") or "")[:16],
            "body_preview": re.sub(r"<\?\s*https?://\S+>?", "[link]", (s.get("body_preview") or ""))[:150],
            "entry_id":     s.get("entry_id", ""),
            "direction":    "sent",
        })
    _api_emails = [
        {"n": i, "direction": e.get("direction", "received"), "subject": e["subject"],
         "from": e["from"], "received": e["received"], "body_preview": e["body_preview"]}
        for i, e in enumerate(_ec)
    ]
    p3_user = (
        f"Today is {today_str}. Tomorrow (next working day) is {tomorrow_str}.\n\n"
        f"EXISTING TASKS:\n{json.dumps(_tsum, indent=1, ensure_ascii=True)}\n\n"
        f"EMAILS (received urgent/needs + sent by Kevin, last 7 days):\n{json.dumps(_api_emails, indent=1, ensure_ascii=True)}"
    )

    # 4. task summaries
    _all_pri = priorities_today + priorities_tomorrow + priorities_week
    _tfs = [
        {"id": e["id"], "title": e["text"], "description": (e.get("description") or "")[:300],
         "actions": e.get("actions", [])[-5:]}
        for e in _all_pri if e.get("id")
    ]
    p4_user = f"Today is {today_str}.\n\nTASKS:\n{json.dumps(_tfs, indent=1, ensure_ascii=True)}"

    # 5. calendar prep
    _cfs = [
        dict(c, prev_meeting_notes=_granola_context.get(f"{c['day']}_{c['idx']}", {}).get("summary", ""))
        for c in _all_day_candidates
    ]
    p5_user = f"Today is {today_str}.\n\nMEETINGS:\n{json.dumps(_cfs, indent=1, ensure_ascii=True)}"

    combined_system = (
        "You are Kevin's inbox / briefing assistant at Oxford University Personnel Services. "
        "This one call performs five independent judgement phases over pre-fetched plain-text "
        "data (no tools, no live fetches). Follow each phase's own verbatim instructions exactly. "
        "Use only plain ASCII punctuation. Return ONLY a single JSON object (no prose, no markdown "
        "fences) with exactly these five top-level keys: context_phase, email_summary_phase, "
        "task_triage_phase, task_summary_phase, calendar_prep_phase -- each holding that phase's "
        "specified output shape."
    )
    _q = '"""'
    combined_user = (
        f"=== 1. context_phase ===\nSystem instructions (verbatim):\n{_q}{p1_sys}{_q}\n{p1_user}\n"
        f'Output shape: {{"context": "...", "subtitle": "..."}}\n\n'
        f"=== 2. email_summary_phase ===\nSystem instructions (verbatim):\n{_q}{_SYS_EMAIL_SUMMARY}{_q}\n{p2_user}\n"
        f'Output shape: {{"<id>": {{"summary": "...", "needs_reply": true/false, "no_action_needed": true/false}}, ...}} keyed by the "id" field above\n\n'
        f"=== 3. task_triage_phase ===\nSystem instructions (verbatim):\n{_q}{_SYS_TRIAGE}{_q}\n{p3_user}\n"
        f'Output shape: {{"new_tasks": [{{"email_n": <n>, "title": "...", "tier": "today|tomorrow|week", "description": "..."}}], "task_updates": [{{"email_n": <n>, "task_id": "...", "note": "..."}}]}}\n\n'
        f"=== 4. task_summary_phase ===\nSystem instructions (verbatim):\n{_q}{_SYS_TASK_SUMMARY}{_q}\n{p4_user}\n"
        f'Output shape: {{"<task id>": "<summary>", ...}}\n\n'
        f"=== 5. calendar_prep_phase ===\nSystem instructions (verbatim):\n{_q}{_SYS_CAL}{_q}\n{p5_user}\n"
        f'Output shape: {{"<day>_<idx>": "2-3 sentences", ...}} where <day>_<idx> is "today_0", "tomorrow_0", "day2_0" etc, matching each meeting\'s "day" and "idx" fields\n\n'
        f"Return the single combined JSON object with all five keys now."
    )

    # Budget: one combined call does all five phases' generation in a single
    # turn. Keep this comfortably under the scheduled task's ExecutionTimeLimit
    # even if the primary account stalls and it fails over to hope@ (2 tries x
    # primary + 1 x fallback).
    t0 = datetime.now()
    try:
        obj = _claude_code_call(combined_system, combined_user, timeout_s=360.0)
    except Exception as e:
        log(f"Phase COMBINED claude_code FAILED (all accounts): {str(e)[:300]}")
        return
    dur = (datetime.now() - t0).total_seconds()
    raw = (obj.get("result") or "").strip()
    if raw.startswith("```"):
        raw = "\n".join(raw.split("\n")[1:])
    if raw.endswith("```"):
        raw = "\n".join(raw.split("\n")[:-1])
    raw = raw.strip()
    _b = raw.find("{")
    if _b > 0:
        raw = raw[_b:]
    try:
        parsed = json.loads(raw)
    except Exception as e:
        log(f"Phase COMBINED parse FAILED: {str(e)[:200]} :: {raw[:300]}")
        return
    _CC_COMBINED = {k: parsed.get(k, {}) for k in
                    ("context_phase", "email_summary_phase", "task_triage_phase",
                     "task_summary_phase", "calendar_prep_phase")}
    _CC_COMBINED_USAGE = obj.get("usage") or {}
    _missing = [k for k in _CC_COMBINED if not parsed.get(k)]
    u = _CC_COMBINED_USAGE
    log(f"Phase COMBINED claude_code OK account={_CC_LAST_ACCOUNT} wall={dur:.1f}s "
        f"in_tok={u.get('input_tokens')} out_tok={u.get('output_tokens')} "
        f"cache_read={u.get('cache_read_input_tokens')} "
        f"cache_creation={u.get('cache_creation_input_tokens')} "
        f"cost_usd={obj.get('total_cost_usd')} missing_keys={_missing or 'none'}")
    try:
        with open(_AI_CALL_LOG, "a", encoding="utf-8") as f:
            f.write(json.dumps({
                "ts": datetime.now().isoformat(timespec="seconds"),
                "seq": "combined", "account": _CC_LAST_ACCOUNT, "model": "claude-haiku-4-5",
                "wall_s": round(dur, 1), "usage": u, "cost_usd": obj.get("total_cost_usd"),
                "num_turns": obj.get("num_turns"), "missing_keys": _missing,
                "phase_output_chars": {k: len(json.dumps(parsed.get(k, ""))) for k in _CC_COMBINED},
            }) + "\n")
    except Exception:
        pass


def _ai_create(model="claude-haiku-4-5", system="", messages=None,
               max_tokens=1024, timeout=None, _phase=None, **_ignored):
    """Drop-in for client.messages.create() at the five triage call sites.
    AI_BACKEND=api        -> the real anthropic client, byte-for-byte as before.
    AI_BACKEND=claude_code -> return this phase's slice of the single combined
                              `claude -p` call (assembled once by
                              _cc_run_combined() before the first phase block)."""
    global _ai_call_seq
    _ai_call_seq += 1
    if AI_BACKEND != "claude_code":
        kw = dict(model=model, system=system, messages=messages, max_tokens=max_tokens)
        if timeout is not None:
            kw["timeout"] = timeout
        return client.messages.create(**kw)

    pk = _PHASE_KEY.get(_phase)
    if pk is not None:
        if _CC_COMBINED is None:
            raise RuntimeError(f"claude_code combined call unavailable (phase={_phase})")
        return _AIText(json.dumps(_CC_COMBINED.get(pk, {}), ensure_ascii=True), _CC_COMBINED_USAGE)

    # No recognised _phase (e.g. a future call site) -> run it standalone.
    user = messages[0]["content"] if messages else ""
    obj = _claude_code_call(system, user, float(timeout or 90.0) + 150.0)
    return _AIText((obj.get("result") or "").strip(), obj.get("usage") or {})

def load_existing_briefing():
    if GITHUB_PAT:
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_PATH}"
            headers = {
                "Authorization": f"token {GITHUB_PAT}",
                "Content-Type":  "application/json",
                "User-Agent":    "work-inbox-script"
            }
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
                data = json.loads(r.read())
            return json.loads(base64.b64decode(data["content"]).decode("utf-8"))
        except Exception as e:
            print(f"WARNING: Could not load existing briefing from GitHub for AI preservation - {e}")
    try:
        local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), GITHUB_PATH)
        with open(local_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"WARNING: Could not load existing briefing locally for AI preservation - {e}")
        return {}

def same_briefing_date(existing, date_label):
    return existing.get("date") == date_label

def cal_summary_key(item):
    return ((item.get("time") or "").strip().lower(),
            (item.get("title") or "").strip().lower())

WEAK_CALENDAR_SUMMARY_PHRASES = [
    "no prep notes available",
    "confirm agenda",
    "confirm scope",
    "no details provided",
    "check calendar context",
    "blocked time or placeholder",
]

def weak_calendar_summary(summary):
    text = (summary or "").strip().lower()
    if not text:
        return True
    if len(text) < 45:
        return True
    return any(phrase in text for phrase in WEAK_CALENDAR_SUMMARY_PHRASES)

def preserve_existing_calendar_summaries(existing, key, items):
    previous = {
        cal_summary_key(item): item.get("summary", "")
        for item in existing.get(key, [])
        if item.get("summary")
    }
    preserved = 0
    for item in items:
        summary = previous.get(cal_summary_key(item))
        if summary:
            current = item.get("summary", "")
            if current and not weak_calendar_summary(current) and weak_calendar_summary(summary):
                continue
            item["summary"] = summary
            preserved += 1
    return preserved

def calendar_summary_count(briefing_doc):
    return sum(
        1
        for key in ("calToday", "calTomorrow", "calDay2", "calDay3")
        for item in briefing_doc.get(key, [])
        if item.get("summary")
    )

def weak_calendar_summary_count(briefing_doc):
    return sum(
        1
        for key in ("calToday", "calTomorrow", "calDay2", "calDay3")
        for item in briefing_doc.get(key, [])
        if item.get("summary") and weak_calendar_summary(item.get("summary"))
    )

def validate_briefing_update(new_doc, old_doc, allow_empty_calendar=False):
    # allow_empty_calendar=True (laptop bridge, WI_BRIDGE_ALLOW_EMPTY_CALENDAR):
    # a run with no calendar source legitimately produces zero calendar summaries
    # and zero calendar-derived absences -- downgrade those three vetoes to
    # warnings. Every other check is unchanged.
    fatal = []
    warnings = []
    context_text = (new_doc.get("context") or "").strip()
    if len(context_text) < 80:
        fatal.append("context is missing or too short")

    if not old_doc:
        return fatal, warnings

    same_day = same_briefing_date(old_doc, new_doc.get("date", ""))
    old_context = (old_doc.get("context") or "").strip()
    if same_day and old_context and len(context_text) < max(80, len(old_context) // 3):
        fatal.append("same-day context would be substantially degraded")

    old_summaries = calendar_summary_count(old_doc)
    new_summaries = calendar_summary_count(new_doc)
    if same_day and old_summaries and new_summaries == 0:
        if allow_empty_calendar:
            warnings.append("same-day calendar summaries removed (bridge mode: no calendar source) - allowed")
        else:
            fatal.append("same-day calendar summaries would be removed")
    elif same_day and old_summaries >= 3 and new_summaries < max(1, old_summaries // 2):
        if allow_empty_calendar:
            warnings.append(f"calendar summaries dropped from {old_summaries} to {new_summaries} (bridge mode) - allowed")
        else:
            fatal.append(f"calendar summaries dropped from {old_summaries} to {new_summaries}")

    old_absences = len(old_doc.get("absences") or [])
    new_absences = len(new_doc.get("absences") or [])
    if same_day and old_absences and new_absences == 0:
        if allow_empty_calendar:
            warnings.append("same-day absences cleared (bridge mode: no calendar source) - allowed")
        else:
            fatal.append("same-day absences would be cleared")

    weak_count = weak_calendar_summary_count(new_doc)
    if weak_count:
        warnings.append(f"{weak_count} calendar summaries look generic or weak")

    return fatal, warnings

def _plain_text(value):
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text.replace(" - ", ": ")

def build_fallback_context(inbox_items, today_items, tomorrow_items):
    unread = [item for item in inbox_items if not item.get("is_read", True)]
    high_importance = [item for item in unread if item.get("importance", 1) == 2]
    sentences = [
        f"Outlook has {len(unread)} unread messages from the last seven days, "
        f"including {len(high_importance)} marked high importance."
    ]

    candidates = high_importance + [item for item in unread if item not in high_importance]
    for item in candidates[:3]:
        sender = _plain_text(item.get("from")) or "an unknown sender"
        subject = _plain_text(item.get("subject")) or "no subject"
        sentences.append(f"Review {subject} from {sender}.")

    def meeting_sentence(label, items):
        timed = [item for item in items if not item.get("all_day")]
        if not timed:
            return f"There are no timed meetings in Outlook {label}."
        names = ", ".join(
            _plain_text(item.get("subject")) or "untitled meeting"
            for item in timed[:3]
        )
        return f"Outlook shows {len(timed)} timed meeting(s) {label}: {names}."

    sentences.append(meeting_sentence("today", today_items))
    sentences.append(meeting_sentence("tomorrow", tomorrow_items))
    sentences.append(
        "AI enrichment is temporarily unavailable, so email cards and calendar data "
        "have been refreshed directly from Outlook without generated interpretation."
    )
    return " ".join(sentences)

def build_fallback_subtitle(inbox_items):
    unread = sum(1 for item in inbox_items if not item.get("is_read", True))
    return f"{unread} unread messages - Outlook-only refresh"

# Two DISTINCT first-call failure modes, handled differently (2026-08-28, Drew):
#
#  1. TRANSIENT BUSY -- pywintypes.com_error (-2147418111, 'Call was rejected
#     by callee.'). Outlook is momentarily busy (mid-sync, a modal dialog
#     open, etc.), not a real fault. Confirmed transient twice on 2026-08-11.
#     -> wait retry_wait_seconds and retry, up to max_attempts. Unchanged.
#
#  2. OUTLOOK NOT RUNNING / NOT CONNECTED -- pywintypes.com_error
#     (-2147352567 / inner -2147221231, 'The file <profile>.ost cannot be
#     accessed. You must connect to Microsoft Exchange at least once...'),
#     an AttributeError on the late-bound Dispatch (COM server still coming
#     up), or CO_E_SERVER_EXEC_FAILURE / RPC-unavailable. Confirmed
#     2026-08-28: a 13:30 reboot left classic OUTLOOK.EXE closed and every
#     scheduled run failed here; the old blind 3x45s wait could never
#     recover it. -> try to LAUNCH classic Outlook once, give it a startup
#     grace period, then retry for real; if it still won't come up (usually
#     an interactive Windows Security / Oxford sign-in prompt only Kevin can
#     clear) fire a SPECIFIC toast telling him exactly what to do.
#
# Retry stays scoped ONLY to this initial connection step -- a real error
# deeper in Phase 1+ still fails immediately instead of being masked.

CLASSIC_OUTLOOK_EXE_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
]
OUTLOOK_STARTUP_GRACE_S = 120

def _launch_classic_outlook():
    """Best-effort start of classic OUTLOOK.EXE. Returns the launched path, or None.

    Launches via explorer.exe so the new Outlook process starts under the
    shell, NOT inside this run's Task Scheduler job object -- that job is
    torn down (killing its child processes) the moment the task ends, so an
    Outlook started as a plain subprocess child would die with the run and
    the next run would have to start it all over again."""
    for exe in CLASSIC_OUTLOOK_EXE_CANDIDATES:
        if os.path.exists(exe):
            try:
                subprocess.Popen(["explorer.exe", exe], close_fds=True)
                return exe
            except Exception as launch_err:
                log(f"Phase 1 - could not launch classic Outlook ({exe}): {launch_err}")
                return None
    log("Phase 1 - classic OUTLOOK.EXE not found in any known location; cannot auto-start it.")
    return None

def _wait_for_outlook_mapi(grace_s):
    """Poll until a fresh COM connect + store mount succeeds, or grace_s elapses."""
    deadline = time.time() + grace_s
    while time.time() < deadline:
        time.sleep(10)
        try:
            probe = win32com.client.dynamic.Dispatch("Outlook.Application")
            probe.GetNamespace("MAPI").GetDefaultFolder(6).Items.Count
            log("Phase 1 - classic Outlook is now MAPI-ready.")
            return True
        except (_COM_ERROR, AttributeError):
            continue
    log(f"Phase 1 - classic Outlook did not become MAPI-ready within {grace_s}s.")
    return False

def _is_outlook_not_ready_error(exc):
    """True if exc means classic Outlook is not running / not MAPI-ready / not
    connected to Exchange -- as opposed to the transient busy-callee case."""
    if isinstance(exc, AttributeError):
        return True
    if isinstance(exc, _COM_ERROR):
        hr    = exc.args[0] if exc.args else None
        inner = exc.args[2] if len(exc.args) > 2 and exc.args[2] else None
        scode = inner[5] if inner and len(inner) > 5 else None
        text  = (inner[2] or "") if inner and len(inner) > 2 else ""
        blob  = f"{exc} {text}".lower()
        if hr in (-2147221231,   # MAPI_E_FAILONEPROVIDER / logon failed
                  -2147221219,   # MAPI_E_NETWORK_ERROR
                  -2146959355,   # CO_E_SERVER_EXEC_FAILURE (Outlook not running, can't auto-start)
                  -2147023174,   # RPC_S_SERVER_UNAVAILABLE
                  -2147417846):  # 'The server threw an exception' / RPC unavailable (alt)
            return True
        if scode in (-2147221231, -2147221219):
            return True
        if (".ost" in blob and "cannot be accessed" in blob) or "connect to microsoft exchange" in blob:
            return True
    return False

def connect_to_outlook(max_attempts=3, retry_wait_seconds=45, allow_launch=True):
    # allow_launch=False: never auto-start OUTLOOK.EXE. That WS1-era auto-start
    # belongs ONLY to the legacy desktop path (MAIL_BACKEND=com). On the laptop
    # (MAIL_BACKEND=imap) a missing/not-connected classic Outlook must degrade
    # to "calendar unavailable this run", never open the app -- that dependency
    # is the whole point of the migration, and Kevin does not want it opening.
    last_error   = None
    tried_launch = False
    for attempt in range(1, max_attempts + 1):
        try:
            # Late binding avoids failures caused by a corrupt win32com.gen_py cache.
            outlook_app  = win32com.client.dynamic.Dispatch("Outlook.Application")
            mapi_ns      = outlook_app.GetNamespace("MAPI")
            # Reach far enough to actually catch a non-mounting Exchange store:
            # touch the folder handle AND force its Items collection to resolve.
            inbox_folder = mapi_ns.GetDefaultFolder(6)
            _ = inbox_folder.Items.Count
            if attempt > 1:
                log(f"Phase 1 - Outlook COM connection succeeded on attempt {attempt}/{max_attempts}.")
            return outlook_app, mapi_ns, inbox_folder
        except (_COM_ERROR, AttributeError) as e:
            last_error = e
            log(f"Phase 1 - Outlook COM connection attempt {attempt}/{max_attempts} failed: {e}")
            if attempt >= max_attempts:
                break
            if _is_outlook_not_ready_error(e):
                if not allow_launch:
                    log("Phase 1 - classic Outlook is not running / not connected "
                        "and allow_launch=False -- NOT starting it. Giving up on COM "
                        "(the caller degrades the calendar to empty).")
                    break
                if not tried_launch:
                    tried_launch = True
                    launched = _launch_classic_outlook()
                    if launched:
                        log(f"Phase 1 - classic Outlook is not running / not connected; started "
                            f"{launched}. Waiting up to {OUTLOOK_STARTUP_GRACE_S}s for MAPI...")
                        _wait_for_outlook_mapi(OUTLOOK_STARTUP_GRACE_S)
                    else:
                        log("Phase 1 - classic Outlook is not running and could not be "
                            "auto-started.")
                        time.sleep(retry_wait_seconds)
                else:
                    log("Phase 1 - classic Outlook still not MAPI-ready after auto-start - "
                        "most likely an interactive Windows Security / Oxford sign-in prompt "
                        "that only Kevin can clear.")
                    time.sleep(retry_wait_seconds)
            else:
                log(f"Phase 1 - Outlook automation layer appears busy (transient). "
                    f"Waiting {retry_wait_seconds}s before retrying...")
                time.sleep(retry_wait_seconds)
    log(f"Phase 1 - Outlook COM connection failed after {max_attempts} attempts. Giving up.")
    if allow_launch and _is_outlook_not_ready_error(last_error):
        _notify_phase_failure(
            "Work Inbox Briefing - Outlook not connected",
            "Classic Outlook is not open / not connected to Exchange. Open "
            r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE, complete any "
            "Windows Security / Oxford sign-in prompt, confirm the status bar reads "
            "'Connected to: Microsoft Exchange' (not Work Offline), then re-run the briefing.")
    raise last_error

if MAIL_BACKEND == "imap":
    # Mail comes from IMAP. Outlook COM is now ONLY a calendar source, and only
    # in the pre-1-Sept interim (CAL_BACKEND=com). Attempt it ONLY when the
    # calendar phases will genuinely use a COM source this run:
    #   - CAL_BACKEND == "com", AND
    #   - not a mail-only parallel capture (WI_MAIL_PARALLEL exits before calendar), AND
    #   - CAL_BACKEND was not requested as 'connector' (Lane B, not built yet).
    # And even then: never auto-launch OUTLOOK.EXE (allow_launch=False). A
    # missing/not-connected classic Outlook degrades to empty calendar + warning.
    _need_com_cal = (CAL_BACKEND == "com" and not MAIL_PARALLEL and not CAL_CONNECTOR_NYI)
    if not _need_com_cal:
        _why = ("WI_MAIL_PARALLEL mail-only capture" if MAIL_PARALLEL
                else "CAL_BACKEND=connector (Lane B not built yet)" if CAL_CONNECTOR_NYI
                else "calendar not in COM scope this run")
        log(f"Phase 1 - MAIL_BACKEND=imap: NOT connecting Outlook COM ({_why}); "
            f"classic Outlook will not be opened. Calendar phases degrade to empty.")
        outlook = mapi = _inbox_folder = None
    else:
        try:
            outlook, mapi, _inbox_folder = connect_to_outlook(allow_launch=False)
        except Exception as _com_e:
            log(f"Phase 1 - MAIL_BACKEND=imap: Outlook COM calendar source unavailable "
                f"({_com_e}); calendar phases degrade to empty, mail briefing continues "
                f"(classic Outlook was NOT launched).")
            outlook = mapi = _inbox_folder = None
else:
    outlook, mapi, _inbox_folder = connect_to_outlook()
cutoff  = datetime.now() - timedelta(days=7)
today   = datetime.now().date()

def next_workday(d):
    d = d + timedelta(days=1)
    while d.weekday() >= 5:
        d = d + timedelta(days=1)
    return d

tomorrow = next_workday(today)

def dt(com_time):
    try:
        return datetime(com_time.year, com_time.month, com_time.day,
                        com_time.hour, com_time.minute, com_time.second)
    except:
        return None

def restrict_date(folder, cutoff_dt):
    """Returns items in `folder` received on/after cutoff_dt.

    Root cause, live-confirmed 12 Aug 2026 (three standalone read-only COM
    diagnostics against the real mailbox, no writes): Outlook COM's
    Items.Restrict() parses the date embedded in the filter string using the
    machine's LOCALE-specific day/month ordering, not the literal field order
    written into the string -- the same underlying class of bug already
    documented for calendar Restrict()+IncludeRecurrences on UK locale (see
    CLAUDE.md "Key Constraints"). The old mm/dd/yyyy-formatted string (e.g.
    '08/05/2026' for 5 Aug) was silently misread as dd/mm (8 May) on this
    UK-locale machine whenever cutoff_dt.day <= 12 -- shifting the real
    cutoff back by months, with Restrict() itself still "succeeding" (no
    exception, a plausible-looking Count). Confirmed directly: for the real
    7-day cutoff on 12 Aug 2026, the mm/dd/yyyy filter returned 562 items
    with the oldest dated 8 May (3+ months old); the dd/mm/yyyy filter for
    the EXACT SAME cutoff returned 63 items, oldest genuinely 5 Aug -- the
    correct number. This -- not "Kevin's inbox is just big" -- is the real
    reason the old >200-item heuristic existed and why it kept firing: a
    misread date bound looks exactly like a big true 7-day window from the
    Count alone, so Count was never a reliable signal either way.

    Fixed at the actual source (dd/mm/yyyy, matching this machine's UK
    locale) rather than patched around with a bigger magic number. A
    defense-in-depth check is kept below in case Restrict() ever silently
    fails for some other reason -- it inspects the actual date of the oldest
    item returned (not the count) to decide whether the filter genuinely
    applied, and the fallback below preserves the date cutoff via bounded
    manual iteration instead of the old behaviour of discarding the cutoff
    entirely and scanning the whole unbounded folder.
    """
    filter_str = "[ReceivedTime] >= '" + cutoff_dt.strftime("%d/%m/%Y %I:%M %p") + "'"
    try:
        restricted = folder.Items.Restrict(filter_str)
        # Verify the date bound genuinely applied by checking the single
        # oldest item actually returned, rather than trusting Count. Uses a
        # separate Restrict() call so sorting this check doesn't disturb the
        # iteration order of `restricted`, which is returned to the caller
        # unchanged on the healthy path.
        check = folder.Items.Restrict(filter_str)
        check.Sort("[ReceivedTime]", False)  # ascending -- oldest first
        oldest = check.GetFirst()
        if oldest is not None:
            oldest_dt = dt(oldest.ReceivedTime)
            # 6-hour grace window absorbs timezone/rounding noise around the
            # boundary itself -- not a loophole for a genuinely stale filter.
            if oldest_dt and oldest_dt < (cutoff_dt - timedelta(hours=6)):
                raise Exception(
                    f"restrict_date: Restrict() returned an item ({oldest_dt}) "
                    f"older than the requested cutoff ({cutoff_dt}) - filter did not apply"
                )
        return restricted
    except Exception:
        # Bounded fallback: manual date-checked iteration, newest first,
        # stopping as soon as an item older than cutoff_dt is reached --
        # preserves the date cutoff instead of the old unrestricted full-
        # folder scan (the actual cause of items up to 4.5 months old
        # reaching FYI). Comparison is done in plain Python datetimes via
        # dt(), so it is immune to the Restrict()-string locale issue above.
        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # descending -- newest first
        bounded = []
        item = items.GetFirst()
        while item is not None:
            try:
                item_dt = dt(item.ReceivedTime)
                if item_dt and item_dt < cutoff_dt:
                    break
                bounded.append(item)
            except Exception:
                pass
            item = items.GetNext()
        return bounded

# -- Phase 1 -- pull Outlook data --
log("Phase 1 - pulling Outlook data...")
inbox = []
unread_count = 0
read_count   = 0
MAX_UNREAD   = 50
MAX_READ     = 30

# VIP senders -- always captured regardless of cap
VIP_NAMES = {
    'Athena Artuso','Marie Cooksey','Sarah Rowles','Simon Burford',
    'Asta Palmer','James Salas Guillen',"Michael O'Sullivan",
    'Anna Carter-Windle','Anthony Kong','Beth Gray','Christopher Sanders',
    'David Johnson','Emma Fitz-Gibbon','Henry Acheampong','Iyanuloluwa Akinsanya',
    'Julie Hickman','Marie King','Michelle Williams','Nathan Kirwan',
    'Susan Pratt','Anne Mortimer','Nicholas Chandler','Steve McBrearty',
}
VIP_EMAILS = {
    'tony.boydell@it.ox.ac.uk','erika.braverman@it.ox.ac.uk',
    'hr.systems@admin.ox.ac.uk','support.access@theaccessgroup.com',
    'edward.demetillo@cority.com','crispin.muncaster@it.ox.ac.uk',
    'christopher.sanders@admin.ox.ac.uk','henry.acheampong@admin.ox.ac.uk',
    'iyanuloluwa.akinsanya@tss.ox.ac.uk',
}

def is_vip(msg):
    try:
        return (msg.SenderName or '').strip() in VIP_NAMES or \
               (msg.SenderEmailAddress or '').lower().strip() in VIP_EMAILS
    except:
        return False

# Kevin's own address, for the to-vs-cc primary-recipient signal (agent-commons
# issue #3 step-3 brief, needs_reply precision fix -- Kevin was cc'd or the
# thread was stale on ~20 of 24 flagged entries in the first real batch,
# because the classifier had no visibility into either dimension before this).
KEVIN_EMAIL = "kevin.lelitte@admin.ox.ac.uk"

# PR_SMTP_ADDRESS: MAPI property tag for a recipient's real SMTP address.
# msg.To / msg.CC return whatever the SENDING client resolved at compose
# time -- for GAL-resolvable recipients (confirmed live 20 Aug 2026: every
# internal @admin.ox.ac.uk sender in a real sample) that's the DISPLAY NAME
# ("Kevin Lelitte; Simon Burford"), not SMTP text, so a substring match
# against KEVIN_EMAIL silently never matches. Live-verified root cause of
# the needs_reply precision regression introduced by this function's first
# version (commit 79c5628f, 10 Aug 2026) -- see begb0037admin/drew memory
# and agent-commons issue #3 for the investigation.
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
olTo = 1

# PR_INTERNET_MESSAGE_ID (Unicode). Captured on each COM mail item ONLY during
# a WI_MAIL_PARALLEL parity capture, so diff_mail_pull.py can join COM<->IMAP
# rows on the stable internet Message-ID (COM EntryID and IMAP UID are
# different namespaces). Outside parity mode this returns "" without touching
# COM at all -> zero change to the live pipeline (no extra PropertyAccessor
# call, and message_id is stripped in inbox_for_api regardless).
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

def _com_message_id(msg):
    if not MAIL_PARALLEL:
        return ""
    try:
        return (msg.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID) or "").strip()
    except Exception:
        return ""

def _kevin_is_primary_recipient(msg):
    """True if Kevin's address appears in To (addressed directly), False if
    only in CC (or not found at all -- distribution lists/aliases mean this
    can't be 100% certain, so it's a signal for the classifier to weigh, not
    an absolute gate on its own).

    Resolves each To-recipient's real SMTP address via PropertyAccessor
    rather than string-matching msg.To, since msg.To can be an
    Exchange-resolved display name (see PR_SMTP_ADDRESS comment above).
    Falls back to the old substring check on msg.To if the Recipients
    collection itself is unavailable, and fails OPEN (True) if both paths
    fail -- never silently suppress a real email over a read failure, same
    philosophy as the original function."""
    try:
        for r in msg.Recipients:
            try:
                if r.Type != olTo:
                    continue
                smtp = r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
                if smtp and smtp.lower() == KEVIN_EMAIL:
                    return True
            except Exception:
                continue  # this recipient couldn't be resolved -- check the rest
        return False
    except Exception:
        try:
            to_field = (msg.To or "").lower()
            return KEVIN_EMAIL in to_field
        except Exception:
            return True  # can't tell -- don't silently suppress a real email over a read failure

# Reuses the folder handle connect_to_outlook() already opened (and retried)
# above, rather than issuing a second unretried GetDefaultFolder(6) call --
# this is exactly the call site both of today's real failures hit.
for msg in ([] if MAIL_BACKEND == "imap" else restrict_date(_inbox_folder, cutoff)):
    try:
        if unread_count >= MAX_UNREAD and read_count >= MAX_READ:
            break
        is_read = not msg.UnRead
        if is_read and read_count >= MAX_READ:
            continue
        if not is_read and unread_count >= MAX_UNREAD:
            continue
        entry = {
            "subject":         msg.Subject,
            "from":            msg.SenderName,
            "from_email":      msg.SenderEmailAddress,
            "received":        str(msg.ReceivedTime),
            "is_read":         is_read,
            "has_attachments": msg.Attachments.Count > 0,
            "importance":      msg.Importance,
            "entry_id":        msg.EntryID,
            "message_id":      _com_message_id(msg),
            "kevin_is_primary_recipient": _kevin_is_primary_recipient(msg)
        }
        if not is_read:
            entry["body_preview"] = (msg.Body or "")[:150]
            unread_count += 1
        else:
            read_count += 1
        inbox.append(entry)
    except:
        continue

inbox.sort(key=lambda x: (not x["is_read"], x["received"]), reverse=True)

# VIP sweep -- pick up any VIP emails missed by the cap
captured_ids = {e["entry_id"] for e in inbox}
for msg in ([] if MAIL_BACKEND == "imap" else restrict_date(mapi.GetDefaultFolder(6), cutoff)):
    try:
        if msg.EntryID in captured_ids:
            continue
        if not is_vip(msg):
            continue
        is_read = not msg.UnRead
        entry = {
            "subject":         msg.Subject,
            "from":            msg.SenderName,
            "from_email":      msg.SenderEmailAddress,
            "received":        str(msg.ReceivedTime),
            "is_read":         is_read,
            "has_attachments": msg.Attachments.Count > 0,
            "importance":      msg.Importance,
            "entry_id":        msg.EntryID,
            "message_id":      _com_message_id(msg),
            "kevin_is_primary_recipient": _kevin_is_primary_recipient(msg)
        }
        if not is_read:
            entry["body_preview"] = (msg.Body or "")[:150]
        inbox.append(entry)
        captured_ids.add(msg.EntryID)
    except:
        continue

inbox.sort(key=lambda x: (not x["is_read"], x["received"]), reverse=True)
print(f"Phase 1 VIP sweep done - total inbox now: {len(inbox)}")

# -- Phase 1c -- named subfolder sweep --
# Kevin confirmed (18 Aug 2026) that Outlook rules auto-file certain mail
# into subfolders under Inbox, and Phase 1 above has only ever read the
# top-level Inbox -- confirmed live the same day: Michael O'Sullivan's
# "RE: Volunteering Leave" reply landed in Inbox/Team/Michael O'Sullivan and
# never reached the briefing (see begb0037admin/drew memory/index.json,
# id starting "2026-08-18-work-inbox-phase-1-pull-only-scans..."). Kevin
# gave explicit scope: recurse into these five named trees only, not the
# whole mailbox. Top-level Inbox pull above is completely unchanged.
#
# Live folder names were verified via COM before hardcoding (18 Aug 2026,
# diag_subfolders.py, read-only) -- two of Kevin's five requested names did
# not match the live folder name exactly:
#   - "Health and Safety" -> no folder by that name exists; the live folder
#     is "H&S" (confirmed as the intended tree -- it's also referenced by
#     the sibling folder "DTP1334 - H&S System Evaluation" under Projects,
#     same abbreviation convention).
#   - "Bi-Monthly CDRPD Working Group" -> was configured as
#     "Bi-monthly CDR/PD working group". REMOVED 29 Aug 2026: Kevin confirmed
#     "I don't have a CDR or PDR folder" -- it no longer exists (deleted/
#     renamed since). Both the COM sweep (WARNING + skip) and the IMAP pull
#     (no LIST match) were silently getting nothing from it. Down to 4 trees.
# "Senior Management", "Team" and "Projects" matched exactly as given.
SUBFOLDER_TREES = [
    "Senior Management",
    "H&S",
    "Team",
    "Projects",
]

# Separate, smaller budget from the top-level Inbox's MAX_UNREAD/MAX_READ
# above so a busy subfolder tree can never displace a top-level Inbox item
# Kevin needs to see -- these two budgets are additive, not shared, and
# both are still bounded by the same 7-day `cutoff` restrict_date() applies
# everywhere else in Phase 1. A live volume check across all 5 trees (18
# Aug 2026) found only 10 items in the last 7 days, so this cap is
# deliberately generous headroom relative to today's real numbers, not a
# tight fit -- documented decision, not a guess.
SUBFOLDER_MAX_UNREAD = 40
SUBFOLDER_MAX_READ   = 20

def walk_folder_tree(folder):
    """Yield `folder` itself, then every nested subfolder, recursively."""
    yield folder
    for sub in folder.Folders:
        yield from walk_folder_tree(sub)

def _build_subfolder_entry(msg, is_read, source_folder):
    entry = {
        "subject":         msg.Subject,
        "from":            msg.SenderName,
        "from_email":      msg.SenderEmailAddress,
        "received":        str(msg.ReceivedTime),
        "is_read":         is_read,
        "has_attachments": msg.Attachments.Count > 0,
        "importance":      msg.Importance,
        "entry_id":        msg.EntryID,
        "message_id":      _com_message_id(msg),
        "kevin_is_primary_recipient": _kevin_is_primary_recipient(msg),
        "source_folder":   source_folder,
    }
    if not is_read:
        entry["body_preview"] = (msg.Body or "")[:150]
    return entry

subfolder_unread = 0
subfolder_read   = 0
subfolder_count  = 0
for tree_name in ([] if MAIL_BACKEND == "imap" else SUBFOLDER_TREES):
    try:
        top_folder = None
        for f in _inbox_folder.Folders:
            if f.Name == tree_name:
                top_folder = f
                break
        if top_folder is None:
            print(f"WARNING: Phase 1c - no Inbox subfolder named {tree_name!r} found - skipping this tree this run (folder may have been renamed/removed)")
            continue
        for sub in walk_folder_tree(top_folder):
            if subfolder_unread >= SUBFOLDER_MAX_UNREAD and subfolder_read >= SUBFOLDER_MAX_READ:
                break
            try:
                for msg in restrict_date(sub, cutoff):
                    try:
                        # olMail (43) only -- subfolders can hold meeting
                        # items/receipts too; filter by Class rather than
                        # letting a bare except conflate "not mail" with a
                        # real bug (see begb0037admin/drew memory/index.json
                        # id starting "2026-08-10-outlook-com-sent-items-
                        # folder-contains-non-mail-items...").
                        if getattr(msg, "Class", None) != 43:
                            continue
                        if msg.EntryID in captured_ids:
                            continue
                        is_read = not msg.UnRead
                        if is_read and subfolder_read >= SUBFOLDER_MAX_READ:
                            continue
                        if not is_read and subfolder_unread >= SUBFOLDER_MAX_UNREAD:
                            continue
                        entry = _build_subfolder_entry(msg, is_read, sub.FolderPath)
                        inbox.append(entry)
                        captured_ids.add(msg.EntryID)
                        subfolder_count += 1
                        if is_read:
                            subfolder_read += 1
                        else:
                            subfolder_unread += 1
                    except Exception:
                        continue
            except Exception as e:
                print(f"WARNING: Phase 1c - failed to scan {sub.FolderPath!r} - {e}")
    except Exception as e:
        print(f"WARNING: Phase 1c - failed to process tree {tree_name!r} - {e}")

inbox.sort(key=lambda x: (not x["is_read"], x["received"]), reverse=True)
print(f"Phase 1c subfolder sweep done - added {subfolder_count} (unread:{subfolder_unread} read:{subfolder_read}) from {len(SUBFOLDER_TREES)} named trees - total inbox now: {len(inbox)}")

sent = []
for msg in ([] if MAIL_BACKEND == "imap" else mapi.GetDefaultFolder(5).Items):
    try:
        t = dt(msg.SentOn)
        if t and t >= cutoff:
            sent.append({
                "subject":      msg.Subject,
                "to":           msg.To,
                "sent":         str(msg.SentOn),
                "body_preview": (msg.Body or "")[:100],
                "entry_id":     msg.EntryID,
                "message_id":   _com_message_id(msg)
            })
    except:
        continue

# -- MAIL_BACKEND=imap: the four COM loops above ran empty; source the mail
#    lists from IMAP+OAuth2 instead. Calendar block below is untouched (COM). --
if MAIL_BACKEND == "imap":
    import imap_mail
    try:
        _imap_res = imap_mail.pull(
            cutoff,
            kevin_email=KEVIN_EMAIL,
            vip_names=VIP_NAMES, vip_emails=VIP_EMAILS,
            subfolder_trees=SUBFOLDER_TREES,
            max_unread=MAX_UNREAD, max_read=MAX_READ,
            sub_max_unread=SUBFOLDER_MAX_UNREAD, sub_max_read=SUBFOLDER_MAX_READ,
            log=log,
        )
    except imap_mail.ImapReauthRequired as _re:
        _rmsg = (f"IMAP mail sign-in expired - run 'Re-auth Work Inbox IMAP' "
                 f"on the Desktop. ({_re})")
        log(f"Phase 1 - {_rmsg}")
        if _imap_reauth_toast_due():
            _notify_phase_failure("Work Inbox Briefing", _rmsg)
        raise SystemExit(1)
    inbox = _imap_res["inbox"]
    sent  = _imap_res["sent"]
    print(f"Phase 1 - IMAP mail pull: inbox {len(inbox)} "
          f"(unread {_imap_res['meta']['inbox_unread']}) sent {len(sent)}")

# -- WI_MAIL_PARALLEL: dump the raw mail lists for diff_mail_pull.py, then EXIT.
#    The parity test compares mail fields only (subject / from / importance /
#    is_read / ...). Everything after this point -- calendar (COM), Granola
#    (Phase 3.7), the combined AI call, card building, push -- is pure latency
#    and failure risk for the parity test and is deliberately skipped. This
#    also removes the classic-Outlook-for-calendar dependency from the parity
#    run. Gated entirely on WI_MAIL_PARALLEL, which only "Run Mail Parity
#    Test.bat" sets -- the live briefing never sets it, so live behaviour
#    (Granola included) is untouched. --
if MAIL_PARALLEL:
    _pdir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "parallel")
    _pfx = "imap" if MAIL_BACKEND == "imap" else "com"
    _dump_ok = False
    try:
        os.makedirs(_pdir, exist_ok=True)
        with open(os.path.join(_pdir, f"{_pfx}_inbox_raw.json"), "w", encoding="utf-8") as _f:
            json.dump(inbox, _f, indent=2, default=str)
        with open(os.path.join(_pdir, f"{_pfx}_sent_raw.json"), "w", encoding="utf-8") as _f:
            json.dump(sent, _f, indent=2, default=str)
        _dump_ok = True
        log(f"Phase 1 - WI_MAIL_PARALLEL wrote {_pfx}_inbox_raw.json "
            f"({len(inbox)} items) / {_pfx}_sent_raw.json ({len(sent)} items)")
    except Exception as _pe:
        log(f"ERROR: WI_MAIL_PARALLEL dump failed - {type(_pe).__name__}: {_pe}")
    log(f"Phase 1 - WI_MAIL_PARALLEL ({_pfx}) mail-only capture done; skipping "
        f"calendar / Granola / AI / push. Exiting {'0' if _dump_ok else '1'}.")
    raise SystemExit(0 if _dump_ok else 1)

# PR_SENDER_NAME (MAPI proptag 0x0C1A001E) -- the display name of
# whoever actually submitted/booked the item. For entries booked by an
# admin/department process on the "People Department - HR Systems"
# calendar, Organizer (and PR_SENT_REPRESENTING_NAME, the same field
# Outlook's Organizer property reads) often shows the placeholder
# department name instead of the real person -- but PR_SENDER_NAME
# reliably still names the real person. Confirmed live 21 Aug 2026
# against every placeholder-organizer entry found in the current
# absence window: "Ant's Annual Leave" -> "Anthony Kong", "Asta -
# Annual Leave" -> "Asta Palmer", "SarahR - A/L" -> "Sarah Rowles", a
# timed "JS - Annual Leave" occurrence -> "James Salas Guillen" --
# general mechanism, not special-cased to any one person. Captured
# here at Phase 1 (while the live COM item is still in hand) and used
# later by the absence-detection organizer-placeholder fallback
# instead of subject-text parsing. Recipients/GlobalAppointmentID were
# also tried live and rejected: Recipients on these placeholder-
# organizer entries contains ONLY the placeholder itself (no real
# attendee), and the two real "Ant's Annual Leave" bookings for the
# same person have completely different GlobalAppointmentIDs (they are
# separate bookings, not occurrences of one series), so a
# series/GlobalAppointmentID lookup would not have connected them.
_PR_SENDER_NAME_TAG = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E"

def _get_pr_sender_name(item):
    try:
        return (item.PropertyAccessor.GetProperty(_PR_SENDER_NAME_TAG) or "").strip()
    except Exception:
        return ""

def _get_is_recurring(item):
    try:
        return bool(item.IsRecurring)
    except Exception:
        return False

week_end = today + timedelta(days=6)
lookback  = today - timedelta(days=30)  # catch multi-day absences spanning today
calendar = []
if mapi is None:
    # MAIL_BACKEND=imap and classic Outlook COM was unavailable this run.
    _cal_items = []
    print("WARNING: Outlook COM unavailable - primary calendar not pulled this run "
          "(MAIL_BACKEND=imap); calendar phases degrade to empty, mail briefing continues")
else:
    _cal_items = mapi.GetDefaultFolder(9).Items
    _cal_items.IncludeRecurrences = True
    _cal_items.Sort("[Start]")
for item in _cal_items:
    try:
        t = dt(item.Start)
        if not t:
            continue
        if t.date() > week_end:
            break
        if t.date() < lookback:
            continue
        calendar.append({
            "subject":      item.Subject,
            "start":        str(item.Start),
            "end":          str(item.End),
            "location":     item.Location,
            "organizer":    item.Organizer,
            "sender_name":  _get_pr_sender_name(item),
            "is_recurring": _get_is_recurring(item),
            "body_preview": (item.Body or "")[:100],
            "all_day":      item.AllDayEvent
        })
    except:
        continue

# Also pull the "People Department - HR Systems" shared calendar -- confirmed
# live, 10 Aug 2026, that it's the department's real leave-tracking calendar
# (279 real dated events with proper Organizer fields, reachable as an "Other
# Calendar" nested under Kevin's own primary mailbox via the same COM
# session). Per Kevin's explicit decision: his own Calendar plus this one are
# the absence source of truth -- if someone's leave isn't in either, he does
# not want it surfaced at all. Wrapped in try/except so a folder-structure
# change (permissions, renaming) degrades to "this calendar contributes
# nothing this run" rather than failing Phase 1 entirely.
hr_calendar_count = 0
try:
    _kevin_store = None
    for _store in mapi.Folders:
        if _store.Name == "kevin.lelitte@admin.ox.ac.uk":
            _kevin_store = _store
            break
    if _kevin_store is not None:
        _hr_cal_folder = _kevin_store.Folders("Calendar").Folders("People Department - HR Systems")
        _hr_items = _hr_cal_folder.Items
        _hr_items.IncludeRecurrences = True
        _hr_items.Sort("[Start]")
        for item in _hr_items:
            try:
                t = dt(item.Start)
                if not t:
                    continue
                if t.date() > week_end:
                    break
                if t.date() < lookback:
                    continue
                calendar.append({
                    "subject":      item.Subject,
                    "start":        str(item.Start),
                    "end":          str(item.End),
                    "location":     item.Location,
                    "organizer":    item.Organizer,
                    "sender_name":  _get_pr_sender_name(item),
                    "is_recurring": _get_is_recurring(item),
                    "body_preview": (item.Body or "")[:100],
                    "all_day":      item.AllDayEvent
                })
                hr_calendar_count += 1
            except:
                continue
    else:
        print("WARNING: 'kevin.lelitte@admin.ox.ac.uk' store not found -- People Department - HR Systems calendar not pulled this run")
except Exception as e:
    print(f"WARNING: People Department - HR Systems calendar pull failed - {e}")

unread_total = sum(1 for m in inbox if not m["is_read"])
print(f"Phase 1 done - inbox:{len(inbox)} (unread:{unread_total}) sent:{len(sent)} calendar:{len(calendar)} (of which HR Systems calendar:{hr_calendar_count})")

# -- Phase 2 -- AI writes context paragraph only --
log("Phase 2 - calling Anthropic API for context...")

now          = datetime.now()
today_str    = now.strftime("%A") + " " + str(now.day) + " " + now.strftime("%B %Y")
tomorrow_str = tomorrow.strftime("%A") + " " + str(tomorrow.day) + " " + tomorrow.strftime("%B %Y")
existing_briefing = load_existing_briefing()

# Calendar day-view now covers 4 rolling working days (today, tomorrow, +2,
# +3) instead of just today/tomorrow -- Kevin's explicit request, 10 Aug
# 2026 ("today, tomorrow, day after that, and day after that... when
# tomorrow comes, it will drop and get Friday"). Day+2/Day+3 use the same
# next_workday() weekend-skipping semantics "tomorrow" already used, for
# consistency -- a Thursday's day+2/day+3 are Monday/Tuesday, not a blank
# Saturday/Sunday.
day2 = next_workday(tomorrow)
day3 = next_workday(day2)

# Leave/absence calendar entries are deliberately excluded from the day-view
# columns -- Kevin's explicit call, same request: "I have the annual leave
# on the sidebar so I don't actually need the annual leave to display in my
# calendar." Duplicates the term list ABSENCE_KEYWORDS uses later in this
# file (defined further down, for the sidebar Absences panel) rather than
# reordering the whole file to share one constant -- keep both lists in sync
# if either changes.
_DAY_VIEW_EXCLUDE_KEYWORDS = [
    "annual leave", "a/l", "on leave", "out of office", "ooo",
    "holiday", "away", "sick leave", "non-working day", "non working day"
]

# "non-working day" / "non working day" -- confirmed live, 12 Aug 2026:
# "Marie K: Non-working day" (People Department - HR Systems calendar, real
# recurring entries incl. 13/14 Aug 2026) slipped through this exclusion and
# showed up in the Tomorrow/Friday day-view columns. This is the SECOND
# occurrence of this exact failure class -- a real leave-calendar phrasing
# variant the keyword list hadn't seen yet -- the first being the bare "AL"
# incident directly below (10 Aug 2026). Both terms added to this list, to
# ABSENCE_KEYWORDS, and to ABSENCE_NOISE (further down) so the item is still
# excluded from day-view but still correctly surfaces on the sidebar
# Absences panel -- verified live as "Marie King - off tomorrow, returns
# Friday 14 August" (Organizer field "Marie King" wins over the
# cleaned-subject fallback "Marie K" here, per the existing organizer-
# priority logic below) -- which is what Kevin wants. If a third phrasing
# variant turns up, see the dedicated Outlook metadata proposal in
# HANDOVER.md before adding a fourth keyword-list entry.
# Bare "AL" abbreviation (no slash) -- confirmed live, 10 Aug 2026: two real
# entries titled exactly "Michael - AL" (7 Aug and 10 Aug) slipped through
# both this day-view exclusion and the sidebar ABSENCE_KEYWORDS check below,
# since the list only matched "a/l" with a slash. Kevin's explicit fix
# request: "i dont want it to show." A plain substring match on "al" would
# false-positive constantly (matches inside "annual", "practical", "Sal",
# etc.), so this is a standalone-word regex instead -- \b requires a
# non-word/word transition immediately either side, which "annual"/
# "practical" never have around their trailing "al" (the preceding letter is
# itself a word character), but "Michael - AL" does (bounded by a space and
# a dash). Applied as an additional OR condition alongside the existing
# substring keyword lists, not a replacement for them.
_BARE_AL_RE = re.compile(r"\bal\b", re.IGNORECASE)

def _has_bare_al(text):
    return bool(_BARE_AL_RE.search(text or ""))

def _is_leave_item(c):
    subj = c.get("subject") or ""
    subj_lower = subj.lower()
    return any(kw in subj_lower for kw in _DAY_VIEW_EXCLUDE_KEYWORDS) or _has_bare_al(subj)

def _cal_for_date(target_date):
    return [
        c for c in calendar
        if datetime.fromisoformat(c["start"]).date() == target_date and not _is_leave_item(c)
    ]

cal_today    = _cal_for_date(today)
cal_tomorrow = _cal_for_date(tomorrow)
cal_day2     = _cal_for_date(day2)
cal_day3     = _cal_for_date(day3)

inbox_for_api = [{k: v for k, v in m.items() if k not in ("entry_id", "message_id")} for m in inbox]

SYSTEM = """You are Kevin's morning inbox briefing assistant at Oxford University Personnel Services.
Your ONLY job is to write the context paragraph. You do not categorise emails. You do not produce cards.
Return ONLY a valid JSON object with exactly two fields - no preamble, no markdown, no code fences.
Use only plain ASCII punctuation: use - instead of dashes, use ' instead of curly quotes.

Return exactly this:
{
  "context": "<A dense, specific 5-7 sentence morning briefing for Kevin. Must include: full names and exact return dates of every absent colleague; which specific projects, systems or cases are blocked because of those absences; any emails waiting more than 48 hours without a response; the most time-critical deadline this week with its exact date; the one thing Kevin should open first. Use real names, real dates, real case numbers and real project names from the data. Every sentence must contain at least one specific proper noun. Do not generalise. Do not mention GitHub, CI/CD, or workflow authentication issues.>",
  "subtitle": "<one short phrase describing the day>"
}"""

USER = f"""Today is {today_str}. Tomorrow (next working day) is {tomorrow_str}.

INBOX ({len(inbox_for_api)} emails, last 7 days):
{json.dumps(inbox_for_api, indent=2, ensure_ascii=True)}

SENT ({len(sent)} items, last 7 days):
{json.dumps(sent, indent=2, ensure_ascii=True)}

CALENDAR TODAY:
{json.dumps(cal_today, indent=2, ensure_ascii=True)}

CALENDAR TOMORROW:
{json.dumps(cal_tomorrow, indent=2, ensure_ascii=True)}
"""

# claude_code backend never touches the anthropic client -- and is typically
# launched with ANTHROPIC_API_KEY unset (subscription billing), which would make
# anthropic.Anthropic() raise at construction. Only build it for the api backend.
client   = anthropic.Anthropic(timeout=60.0) if AI_BACKEND == "api" else None
anthropic_available = True
context  = ""
subtitle = ""
if AI_BACKEND == "claude_code":
    # Phase 2's context is produced by the ONE combined `claude -p` call,
    # assembled + fired just after Phase 3 card-building (the earliest point
    # all five phase payloads exist). Parsed there via _p2_finalise().
    print("Phase 2 - deferred to the single combined claude_code call")
else:
    try:
        response = _ai_create(
            model      = "claude-haiku-4-5",
            max_tokens = 1024,
            system     = SYSTEM,
            messages   = [{"role": "user", "content": USER}],
            _phase     = "context",
        )

        raw_text = response.content[0].text.strip()
        if raw_text.startswith("```"):
            raw_text = "\n".join(raw_text.split("\n")[1:])
        if raw_text.endswith("```"):
            raw_text = "\n".join(raw_text.split("\n")[:-1])

        ai_output = json.loads(raw_text)
        context  = ai_output.get("context", "")
        subtitle = ai_output.get("subtitle", "")
    except Exception as e:
        anthropic_available = False
        print(f"WARNING: Phase 2 context failed - {e}")
    context, subtitle = _p2_finalise(context, subtitle)
    print("Phase 2 done - context written")

# -- Phase 3 -- Python builds every card --
log("Phase 3 - building cards from inbox...")

# Categorisation rules -- applied in order, first match wins
# importance: 0=low, 1=normal, 2=high
URGENT_SENDERS   = []  # add sender email fragments here if needed
URGENT_SUBJECTS  = ["major incident", "priority 1", "p1", "urgent", "critical", "security vulnerab"]
NEEDS_SUBJECTS   = ["re:", "fw:", "fwd:", "action", "required", "please", "timeline", "update",
                    "chasing", "waiting", "overdue", "follow", "scoping", "handover", "error",
                    "import", "failed", "issue", "case ", "support"]
FYI_SUBJECTS     = ["fyi", "notification", "scheduled", "maintenance", "summary", "workshop",
                    "invitation", "invite", "digest", "recap", "newsletter", "annual leave",
                    "out of office", "automatic reply", "accepted:", "declined:", "cancelled:"]
LOW_SUBJECTS     = ["unsubscribe", "noreply", "no-reply", "do not reply", "automated",
                    "github", "pages", "build", "deploy", "run failed", "wisp"]

def categorise(msg):
    subj    = (msg.get("subject") or "").lower()
    sender  = (msg.get("from_email") or "").lower()
    is_read = msg.get("is_read", True)
    imp     = msg.get("importance", 1)

    # High importance flag always pushes to urgent
    if imp == 2:
        return "urgent"

    # Subject keyword matching
    for kw in LOW_SUBJECTS:
        if kw in subj or kw in sender:
            return "low"
    for kw in URGENT_SUBJECTS:
        if kw in subj:
            return "urgent"
    # Unread + needs keywords -- needs response
    if not is_read:
        for kw in NEEDS_SUBJECTS:
            if kw in subj:
                return "needs"
    for kw in FYI_SUBJECTS:
        if kw in subj:
            return "fyi"
    # Unread with no other match -- needs response
    if not is_read:
        return "needs"
    # Read with no match -- fyi
    return "fyi"

def badge_for(msg, category):
    imp  = msg.get("importance", 1)
    received = msg.get("received", "")
    age_hrs = 0
    try:
        t = datetime.fromisoformat(received.split("+")[0].split(" (")[0].strip())
        age_hrs = (datetime.now() - t).total_seconds() / 3600
    except:
        pass

    if category == "urgent":
        return "Act today", "red"
    if category == "needs":
        if age_hrs > 48:
            return "Overdue", "red"
        return "Reply within 48hrs", "yellow"
    if category == "fyi":
        return "FYI", "gray"
    return "", "gray"

def make_card(msg, category):
    subj    = msg.get("subject") or "(no subject)"
    sender  = msg.get("from") or ""
    preview = (msg.get("body_preview") or "").strip()
    preview = re.sub(r"<\?\s*https?://\S+>?", "[link]", preview)
    badge, badge_type = badge_for(msg, category)

    title = subj
    sub   = f"From <strong>{sender}</strong>."
    if preview:
        sub += f" {html.escape(preview[:120])}"

    received_str = ""
    try:
        rec = msg.get("received", "")
        rec_dt = datetime.fromisoformat(rec.split("+")[0].split(" (")[0].strip())
        received_str = str(rec_dt.day) + rec_dt.strftime(" %b")
    except:
        pass
    card = {
        "title":     title,
        "sub":       sub,
        "badge":     badge,
        "badgeType": badge_type,
        "subject":   subj,
        "from":      sender,
        "entry_id":  msg.get("entry_id", ""),
        "received":  received_str,
        "received_raw": msg.get("received", ""),
        "kevin_is_primary_recipient": msg.get("kevin_is_primary_recipient", True)
    }
    return card

urgent = []
needs  = []
fyi    = []
low    = []

for msg in inbox:
    cat  = categorise(msg)
    card = make_card(msg, cat)
    if cat == "urgent":
        urgent.append(card)
    elif cat == "needs":
        needs.append(card)
    elif cat == "fyi":
        fyi.append(card)
    else:
        low.append(card)

print(f"Phase 3 done - urgent:{len(urgent)} needs:{len(needs)} fyi:{len(fyi)} low:{len(low)}")

# -- Phase 3.2 - AI summaries for urgent/needs email cards --
# Same pattern as Phase 3.7's priority-task summaries, applied to raw email
# cards instead - so Urgent/Needs show a genuine one-sentence summary rather
# than the first ~150 characters of the email body verbatim (card["sub"]).
log("Phase 3.2 - generating AI email summaries...")
summary_candidates = [c for c in (urgent + needs) if c.get("entry_id")]
# Entry IDs of cards actually demoted by Phase 3.3/3.3b below (AI-confirmed
# no_action_needed). Declared unconditionally, before the Phase 3.2 block
# below, so it always exists (empty set) even if Phase 3.2 is skipped or
# fails outright. Reused by Phase 3.5's Command Centre task-suggestion
# triage further down so a bogus new_tasks suggestion isn't generated for an
# email this same run already confirmed Kevin has nothing to do about --
# see that phase's own comment for the full reasoning and its own scope
# limits (12 Aug 2026, extending the same-day Phase 3.3 Needs fix).
_noise_demoted_entry_ids = set()

# ---------------------------------------------------------------------------
# claude_code backend: assemble + fire the SINGLE combined `claude -p` call
# now. All five phases (context / email summaries / task triage / task
# summaries / calendar prep) go in ONE prompt; each phase block below then
# reads its slice from _CC_COMBINED via _ai_create(_phase=...). The api
# backend is untouched -- it keeps five separate client.messages.create()
# calls. See docs/CLAUDE_CODE_BACKEND.md / docs/COLLAPSE_TO_ONE_CALL_PLAN.md.
# ---------------------------------------------------------------------------
if AI_BACKEND == "claude_code":
    _cc_load_priorities()               # cc_content + priorities_{today,tomorrow,week}
    _cc_build_cal_candidates_early()    # _all_day_candidates from raw per-day calendars
    _cc_fetch_granola(_all_day_candidates)   # _granola_context
    _cc_run_combined()                  # -> _CC_COMBINED (or None on failure)
    _p2s = (_CC_COMBINED or {}).get("context_phase") or {}
    context  = _p2s.get("context", "")  if isinstance(_p2s, dict) else ""
    subtitle = _p2s.get("subtitle", "") if isinstance(_p2s, dict) else ""
    context, subtitle = _p2_finalise(context, subtitle)
    if _CC_COMBINED is None:
        anthropic_available = False
        print("Phase 2 - combined claude_code call FAILED; downstream AI phases will skip")
    else:
        print("Phase 2 done - context written (combined claude -p call)")

if summary_candidates and anthropic_available:
    try:
        # Use short sequential ids in the API exchange, not the raw ~140-char
        # Outlook EntryID -- confirmed live, 10 Aug 2026: with 157 real
        # urgent+needs candidates, using entry_id as the JSON key hit
        # stop_reason="max_tokens" at the full 8000-token budget (hex strings
        # tokenize far less efficiently than English text, so the KEYS alone
        # were consuming most of the budget before the model reached the
        # actual summaries). Switching to "0","1","2"... as the wire-format
        # id and mapping back to entry_id locally resolved it completely on
        # the identical real payload: stop_reason="end_turn", only 5947/8000
        # tokens used, all 157 entries parsed. Root cause was token-inefficient
        # keys, not response size -- raising max_tokens further would not
        # have fixed this on its own.
        # Recipient-role + age signals -- confirmed root cause, 10 Aug 2026:
        # Lauren's review of the first real needs_reply batch found ~20 of 24
        # flagged entries were cc-only threads or clearly stale, and neither
        # signal reached the classifier before this fix (no To/CC captured at
        # all for received mail, no date passed in this payload even though
        # it's one field away). Both now computed deterministically in Python
        # and given to the model as explicit signals, not left for it to guess.
        def _age_days(card):
            try:
                rec_dt = datetime.fromisoformat(card.get("received_raw", "").split("+")[0].split(" (")[0].strip())
                return (datetime.now() - rec_dt).days
            except:
                return None

        emails_for_summary = [
            {
                "id":      str(i),
                "subject": c["subject"],
                "from":    c["from"],
                "preview": (c.get("sub") or "")[:250],
                "kevin_is_primary_recipient": c.get("kevin_is_primary_recipient", True),
                "age_days": _age_days(c)
            }
            for i, c in enumerate(summary_candidates)
        ]
        EMAIL_SUMMARY_SYSTEM = _SYS_EMAIL_SUMMARY  # verbatim; hoisted to module scope
        email_summary_user = (
            f"Today is {today_str}.\n\n"
            f"EMAILS:\n{json.dumps(emails_for_summary, indent=1, ensure_ascii=True)}"
        )
        es_resp = _ai_create(
            model      = "claude-haiku-4-5",
            # Raised 8000 -> 14000, 12 Aug 2026: the 10 Aug incident that
            # first hit stop_reason=max_tokens on this call used 5947/8000
            # for 157 entries x 2 fields (summary, needs_reply). Adding the
            # no_action_needed field (Phase 3.3 demotion work, same day)
            # brings this call to a real live inbox size of 165 entries x 3
            # fields -- flagged by Codex review as leaving uncomfortably
            # little headroom against a repeat of that exact failure mode,
            # not yet re-tested against a real payload at this size before
            # this change, so erring toward more headroom rather than
            # waiting to find out live.
            max_tokens = 14000,
            # Per-call override, 12 Aug 2026: raising max_tokens above
            # without also giving the call more wall-clock time doesn't
            # actually help -- confirmed live the same session, this exact
            # call failed with "Request timed out or interrupted" against
            # the global `client = anthropic.Anthropic(timeout=60.0)`
            # default (line ~551) immediately after the max_tokens increase
            # above, on a real 165-entry payload. Scoped to this one call
            # only (not raising the global client timeout) since this is by
            # far the largest-volume/longest call in the file and the other
            # client.messages.create() call sites haven't shown this issue.
            timeout    = 150.0,
            system     = EMAIL_SUMMARY_SYSTEM,
            messages   = [{"role": "user", "content": email_summary_user}],
            _phase     = "email_summary",
        )
        es_raw = es_resp.content[0].text.strip()
        if es_raw.startswith("```"):
            es_raw = "\n".join(es_raw.split("\n")[1:])
        if es_raw.endswith("```"):
            es_raw = "\n".join(es_raw.split("\n")[:-1])
        email_summaries = json.loads(es_raw)
        applied = 0
        needs_reply_count = 0
        # Deterministic staleness gate -- Kevin's explicit cutoff. Set to 2
        # months/60 days (10 Aug 2026), briefly revised to 30 days (11 Aug
        # 2026), then reverted back to 60 the same day -- 60 is his current
        # final call on this parameter. This is a hard override applied
        # AFTER the AI's own judgement, not a replacement for it -- the AI
        # still gets age_days as a soft signal above (for anything younger
        # than the cutoff), but nothing older than 60 days can end up
        # needs_reply=true regardless of what the model decides, since
        # Kevin was explicit that this is his call to set, not the
        # model's to infer.
        STALENESS_CUTOFF_DAYS = 60
        stale_overridden = 0
        for i, c in enumerate(summary_candidates):
            entry = email_summaries.get(str(i))
            if entry is None:
                continue
            # Defensive: accept either the new {summary, needs_reply,
            # no_action_needed} shape or, if the model ever reverts to a
            # bare string, treat that as needs_reply/no_action_needed both
            # defaulting to False rather than crashing the phase.
            # Track whether this was a genuine, schema-valid AI verdict (a
            # real dict response with actual booleans for BOTH needs_reply
            # and no_action_needed) versus one of the defensive fallback
            # paths above (bare-string response, or a dict missing/
            # mistyping either field) -- fallbacks also land on False, but
            # that's a "we don't actually know" default, not a real AI
            # judgement. Phase 3.3 below must only ever act on the genuine
            # case -- see its own comment for why this distinction matters.
            # Also reject a contradictory model verdict (needs_reply=true
            # AND no_action_needed=true at the same time -- the prompt tells
            # the model never to do this, but nothing enforces it) as
            # invalid too. Checked against the RAW entry here, before the
            # staleness override below can run -- that override only ever
            # flips needs_reply true->false, which would otherwise turn an
            # already-contradictory {true, true} pair into a {false, true}
            # pair that looks like a clean demotion candidate despite never
            # having been a coherent verdict to begin with.
            c["_ai_verdict_valid"] = (
                isinstance(entry, dict)
                and isinstance(entry.get("needs_reply"), bool)
                and isinstance(entry.get("no_action_needed"), bool)
                and not (entry.get("needs_reply") is True and entry.get("no_action_needed") is True)
            )

            if isinstance(entry, dict):
                c["ai_summary"] = entry.get("summary", "")
                nr = entry.get("needs_reply", False)
                c["needs_reply"] = bool(nr) if isinstance(nr, bool) else False
                na = entry.get("no_action_needed", False)
                c["no_action_needed"] = bool(na) if isinstance(na, bool) else False
            else:
                c["ai_summary"] = str(entry)
                c["needs_reply"] = False
                c["no_action_needed"] = False

            age = _age_days(c)
            if c["needs_reply"] and age is not None and age > STALENESS_CUTOFF_DAYS:
                c["needs_reply"] = False
                stale_overridden += 1
                # Staleness forces needs_reply true->false as a deterministic
                # override, but says nothing about whether Kevin still has
                # some OTHER action to take -- an old, stale-flagged thread
                # isn't necessarily a pure FYI. Leave no_action_needed as
                # whatever the model itself said, don't infer it from this.

            applied += 1
            if c["needs_reply"]:
                needs_reply_count += 1
        print(f"Phase 3.2 done - {applied} email summaries generated, {needs_reply_count} flagged needs_reply ({stale_overridden} overridden false for being older than {STALENESS_CUTOFF_DAYS} days)")

        # -- Phase 3.3 -- demote AI-confirmed no-action cards out of Needs --
        # Kevin's explicit request, 12 Aug 2026: Phase 3's categorise() (the
        # urgent/needs/fyi/low split above) runs BEFORE any AI involvement,
        # on subject-keyword + read/unread rules alone ("re:", "chasing",
        # "follow", etc.) -- it has no idea whether Kevin himself needs to
        # act, only whether the subject looks actionable. That's why
        # colleague-to-colleague threads he's only cc'd on land in Needs by
        # keyword match, even though this same Phase 3.2 AI pass -- reading
        # the actual content a few seconds later -- correctly judges most of
        # them needs_reply=false.
        #
        # REVISION, same session: the first version of this demotion used
        # needs_reply=False AND the model's summary text literally
        # containing "no action needed" as a second safety check. That text
        # heuristic looked solid against one live snapshot (98/108 of that
        # run's needs_reply=false Needs cards used that exact phrase) but
        # failed completely on the very next live run -- the model's
        # freeform wording is not deterministic between calls, and a
        # same-content re-run produced 0/108 matches (different phrasing,
        # e.g. "Kevin is cc'd only" instead of "no action needed"). Relying
        # on exact-text matching against non-deterministic LLM prose is the
        # same brittleness class as the subject-keyword-list gap fixed
        # earlier today (Marie K "non-working day") -- chasing wording
        # variants is a losing game. Replaced with a real structured signal
        # instead: the model now returns an explicit `no_action_needed`
        # boolean alongside `needs_reply` (see EMAIL_SUMMARY_SYSTEM above),
        # not inferred from prose at all.
        #
        # Deliberately conservative on one remaining axis, checked against a
        # Codex read-only review before building (12 Aug 2026):
        # Requires `_ai_verdict_valid` is True -- i.e. the model returned a
        # genuine dict response with real booleans for BOTH needs_reply and
        # no_action_needed, not one of the defensive fallback paths above
        # (bare-string response, missing/mistyped field). Those fallbacks
        # default both to False as a "we don't know" case, not a real
        # verdict, and must never drive a demotion.
        #
        # EXTENSION, same day (12 Aug 2026): originally this only demoted
        # from Needs -- Urgent (importance-flagged or urgent-keyword-matched
        # mail) was left alone as a materially higher-risk call, and Phase
        # 3.5's separate Command Centre task-suggestion triage was flagged
        # as a known, unfixed gap (it independently re-derives its own
        # candidate list via a fresh categorise(m) call on raw inbox
        # messages and had no needs_reply/no_action_needed field to consult
        # at all). Kevin approved extending both after ~9 similarly-noisy
        # Urgent cards were seen live in the original session. See Phase
        # 3.3b immediately below for the Urgent extension, and Phase 3.5's
        # own comment further down for how it reuses `_noise_demoted_entry_ids`
        # (built by both this block and 3.3b) rather than re-deriving noise
        # detection from scratch.
        # Exception-safety, added after a Codex review of this exact block
        # caught a real bug: an earlier version mutated `needs`/`fyi`
        # card-by-card DURING the loop and only reassigned `needs =
        # still_needs` at the end, so an exception partway through (e.g. a
        # non-string received_raw breaking the later sort) would leave some
        # cards already appended to `fyi` while `needs` still held them too
        # -- a real duplicate-card bug -- and would also skip the internal-
        # field cleanup below, leaking it into the public briefing.json.
        # Fixed by building into local temp lists first and only committing
        # to `needs`/`fyi` after the whole pass succeeds, wrapping the whole
        # thing in its own try/except so a failure here can never take down
        # Phase 3.2's already-good results, and moving cleanup into
        # `finally` so it always runs regardless of outcome.
        try:
            demoted_count = 0
            still_needs = []
            newly_fyi = []
            demoted_ids_this_pass = set()
            for card in needs:
                if card.get("_ai_verdict_valid") and card.get("needs_reply") is False and card.get("no_action_needed") is True:
                    card["badge"], card["badgeType"] = badge_for(card, "fyi")
                    newly_fyi.append(card)
                    demoted_count += 1
                    eid = card.get("entry_id")
                    if eid:
                        demoted_ids_this_pass.add(eid)
                else:
                    still_needs.append(card)
            needs = still_needs
            fyi.extend(newly_fyi)
            # Only merge into the shared cross-phase set once the tier/FYI
            # lists above are actually committed (i.e. we know this pass
            # genuinely succeeded) -- Codex review flagged that adding to a
            # shared set mid-loop, before commit, could suppress a Phase 3.5
            # task suggestion for a card that a later exception in THIS pass
            # then left un-demoted after all.
            _noise_demoted_entry_ids.update(demoted_ids_this_pass)
            if demoted_count:
                try:
                    fyi.sort(key=lambda c: str(c.get("received_raw") or ""), reverse=True)
                except Exception as sort_err:
                    print(f"WARNING: Phase 3.3 FYI re-sort failed, order preserved as-is - {sort_err}")
                print(f"Phase 3.3 done - {demoted_count} Needs card(s) demoted to FYI (AI-confirmed no action needed)")
        except Exception as demote_err:
            print(f"WARNING: Phase 3.3 demotion failed, Needs left unchanged - {demote_err}")

        # -- Phase 3.3b -- same demotion logic applied to Urgent, 12 Aug 2026 --
        # Kevin approved extending Phase 3.3 above to Urgent after ~9
        # similarly-noisy Urgent cards were seen live in the original
        # session (importance-flagged or urgent-keyword-matched mail from
        # colleague threads Kevin is only cc'd on). Urgent cards already
        # carry the same AI verdict fields as Needs cards -- summary_candidates
        # above is `urgent + needs` together, so every Urgent card that got
        # an AI summary already has needs_reply/no_action_needed/
        # _ai_verdict_valid set by the exact same loop that set them for
        # Needs cards. No new AI call, same criteria, same exception-safety
        # pattern (atomic temp lists, its own try/except so a failure here
        # can't take down Phase 3.3's already-committed Needs result).
        try:
            demoted_urgent_count = 0
            still_urgent = []
            newly_fyi_from_urgent = []
            demoted_urgent_ids_this_pass = set()
            for card in urgent:
                if card.get("_ai_verdict_valid") and card.get("needs_reply") is False and card.get("no_action_needed") is True:
                    card["badge"], card["badgeType"] = badge_for(card, "fyi")
                    newly_fyi_from_urgent.append(card)
                    demoted_urgent_count += 1
                    eid = card.get("entry_id")
                    if eid:
                        demoted_urgent_ids_this_pass.add(eid)
                else:
                    still_urgent.append(card)
            urgent = still_urgent
            fyi.extend(newly_fyi_from_urgent)
            _noise_demoted_entry_ids.update(demoted_urgent_ids_this_pass)
            if demoted_urgent_count:
                try:
                    fyi.sort(key=lambda c: str(c.get("received_raw") or ""), reverse=True)
                except Exception as sort_err:
                    print(f"WARNING: Phase 3.3b FYI re-sort failed, order preserved as-is - {sort_err}")
                print(f"Phase 3.3b done - {demoted_urgent_count} Urgent card(s) demoted to FYI (AI-confirmed no action needed)")
        except Exception as demote_err:
            print(f"WARNING: Phase 3.3b demotion failed, Urgent left unchanged - {demote_err}")
        finally:
            for card in (urgent + needs + fyi + low):
                card.pop("_ai_verdict_valid", None)
    except Exception as e:
        print(f"WARNING: Phase 3.2 AI email summaries failed - {e}")
elif summary_candidates:
    print("Phase 3.2 skipped - Anthropic is unavailable")
else:
    print("Phase 3.2 skipped - no urgent/needs emails")

# -- Phase 3.3c -- FYI thread-collapse + explicit aging, 12 Aug 2026 --
# Kevin approved cleanup for two more documented FYI/Parked symptoms once the
# true root cause (the restrict_date() locale bug + unbounded VIP sweep
# above) was fixed at the source: (1) 47% of the pre-existing FYI baseline
# was duplicate "RE:"/"FW:" threads with no thread-collapsing anywhere in
# the pipeline (documented live example: "RE: HR Systems Managers Meeting"
# x8); (2) cards landing in FYI via the Phase 3.3/3.3b demotion logic above
# had no downstream cleanup, so they could accumulate indefinitely if a
# future regression ever reopened the date-bound bug fixed above. Placed
# outside the `if summary_candidates and anthropic_available:` block above
# so it always runs -- thread duplication and cutoff-based aging are both
# real regardless of whether the AI summary/demotion phases ran this time.
fyi_raw_count = len(fyi)
try:
    # (1) Thread/subject dedup. Normalizes by repeatedly stripping leading
    # Re:/Fw:/Fwd: prefixes (handles "Re: Fw: ..." chains, case-insensitive)
    # and collapsing whitespace/case, then keeps a single card per thread --
    # the most recently received one -- with an explicit messageCount so the
    # collapse is visible to the dashboard rather than a silent reduction
    # (see the corresponding js/app.js change, same session).
    _RE_FWD_PREFIX = re.compile(r'^\s*(re|fw|fwd)\s*:\s*', re.IGNORECASE)

    def _thread_key(card):
        s = (card.get("subject") or "").strip().lower()
        s = re.sub(r'\s+', ' ', s)
        while True:
            new_s = _RE_FWD_PREFIX.sub('', s).strip()
            if new_s == s:
                break
            s = new_s
        return s or ("id:" + str(card.get("entry_id", "")))

    threads = {}
    thread_order = []
    for card in fyi:
        key = _thread_key(card)
        if key not in threads:
            card["messageCount"] = 1
            threads[key] = card
            thread_order.append(key)
        else:
            existing = threads[key]
            new_count = existing.get("messageCount", 1) + 1
            try:
                is_newer = str(card.get("received_raw") or "") > str(existing.get("received_raw") or "")
            except Exception:
                is_newer = False
            if is_newer:
                card["messageCount"] = new_count
                threads[key] = card
            else:
                existing["messageCount"] = new_count

    fyi = [threads[k] for k in thread_order]
    collapsed_count = fyi_raw_count - len(fyi)

    # (2) Explicit age cutoff, belt-and-braces on top of the restrict_date()
    # fix above (not a replacement for it). Consistent with the pipeline's
    # own existing precedent of date-bounding again at the point of use
    # rather than trusting an upstream pull to stay bounded forever (see
    # STALENESS_CUTOFF_DAYS elsewhere in this file, and Lauren's 60-day
    # drafting-age cutoff in the sibling meeting-records pipeline). If the
    # date-bound fix above ever regresses, this still stops FYI from quietly
    # accumulating months-old cards with nothing removing them.
    FYI_MAX_AGE_DAYS = 7
    _fyi_age_cutoff = datetime.now() - timedelta(days=FYI_MAX_AGE_DAYS)
    _fyi_before_age_filter = len(fyi)

    def _fyi_card_recent_enough(card):
        try:
            rd = str(card.get("received_raw") or "")
            t = datetime.fromisoformat(rd.split("+")[0].split(" (")[0].strip())
            return t >= _fyi_age_cutoff
        except Exception:
            return True  # can't parse -- don't silently drop a real card

    fyi = [c for c in fyi if _fyi_card_recent_enough(c)]
    aged_out_count = _fyi_before_age_filter - len(fyi)

    print(f"Phase 3.3c done - FYI thread-collapse: {fyi_raw_count} raw -> {len(fyi)} threads "
          f"({collapsed_count} collapsed), {aged_out_count} aged out (>{FYI_MAX_AGE_DAYS}d)")
except Exception as fyi_clean_err:
    print(f"WARNING: Phase 3.3c FYI thread-collapse/aging failed, FYI left unchanged - {fyi_clean_err}")

# -- Calendar post-processing --
KNOWN_ABSENCES = []

def build_cal_items(items):
    result = []
    items = sorted(items, key=lambda x: x.get("start", ""))
    for item in items:
        start = item.get("start", "")
        try:
            t = datetime.fromisoformat(start)
            time_str = "All day" if item.get("all_day") else t.strftime("%H:%M")
        except:
            time_str = "All day" if item.get("all_day") else ""

        title = item.get("subject", "")
        sub   = item.get("organizer", "") or ""
        alert = ""

        # Check known absences
        title_lower = title.lower()
        for absence in KNOWN_ABSENCES:
            if any(tr in title_lower for tr in absence["triggers"]):
                time_str = "All day"
                title    = absence["title"]
                sub      = absence["sub"]
                alert    = absence["alert"]
                break

        cal_item = {"time": time_str, "title": title, "sub": sub}
        if alert:
            cal_item["alert"] = alert
        result.append(cal_item)
    return result

# Detect absences from calendar sources only -- Kevin's Calendar plus the
# "People Department - HR Systems" calendar pulled in Phase 1. Explicit
# decision, 10 Aug 2026: these two calendars are the absence source of
# truth; if someone's leave isn't logged in either, Kevin does not want it
# surfaced at all. The previous OOO-auto-reply-email fallback (and the
# best-effort date-guessing built for it the same day) is deliberately
# removed, not just unused -- it was the source of both the "date unknown"
# entries and cross-department noise (e.g. IT Services staff who were never
# going to appear in a People Department leave calendar).
ABSENCE_KEYWORDS = [
    "annual leave", "a/l", "on leave", "out of office", "ooo",
    "holiday", "away", "sick leave", "non-working day", "non working day"
]
ABSENCE_NOISE = [
    "annual leave", "a/l", "on leave", "out of office", "ooo",
    "holiday", "away", "sick leave", "leave", "non-working day",
    "non working day"
]
absence_map = {}

def _parse_iso_date(value):
    try:
        return datetime.fromisoformat(str(value)).date()
    except:
        return None

def _fmt_absence_date(d):
    return d.strftime("%A ") + str(d.day) + d.strftime(" %B")

def _absence_key(name):
    return " ".join((name or "").lower().split())

def _clean_absence_name(text):
    name = (text or "").strip()
    if "<" in name:
        name = name.split("<", 1)[0].strip()
    for sep in [" - ", " -- ", ":", "|"]:
        parts = [p.strip() for p in name.split(sep) if p.strip()]
        if len(parts) == 2:
            left_l = parts[0].lower()
            right_l = parts[1].lower()
            if any(kw in left_l for kw in ABSENCE_NOISE):
                name = parts[1]
            elif any(kw in right_l for kw in ABSENCE_NOISE):
                name = parts[0]
            break
    lower_name = name.lower()
    for kw in ABSENCE_NOISE:
        lower_name = lower_name.replace(kw, " ")
    # Strip a standalone "AL" token too (e.g. "Michael - AL" -> "Michael")
    # -- not added to ABSENCE_NOISE's plain substring loop above because a
    # blind "al" substring replace would corrupt real names that merely
    # contain "al" (Alan, Alison, Natalie, Malcolm...). _BARE_AL_RE only
    # matches "al" as its own word (see the comment above its definition),
    # so "Alan Smith" is untouched while "Michael - AL" -> "Michael - ".
    lower_name = _BARE_AL_RE.sub(" ", lower_name)
    cleaned = []
    for token in lower_name.replace("(", " ").replace(")", " ").split():
        if token in ("is", "on", "until", "from", "to", "for"):
            continue
        cleaned.append(token)
    if cleaned:
        name = _title_case_name(" ".join(cleaned))
    return " ".join(name.split()).strip(" -:")

# str.title() capitalises the letter after ANY non-alpha character,
# which wrongly turns a possessive "'s" suffix into "'S" -- confirmed
# live, 21 Aug 2026: "ant's" (subject-derived token from "Ant's Annual
# Leave") -> "Ant'S" instead of "Ant's". A trailing "'s" token (the
# whole word ends in apostrophe-s, e.g. "ant's", "kevin's") is always
# possessive, never a mid-name apostrophe, so it's handled specially;
# genuine mid-name apostrophes (O'Sullivan, O'Brien) don't end in "'s"
# and still get the standard title()-style capitalise-after-apostrophe
# behaviour, unchanged from before.
def _title_case_name(name):
    out = []
    for token in name.split():
        if len(token) > 2 and token.lower().endswith("'s"):
            base = token[:-2]
            out.append(base[:1].upper() + base[1:].lower() + "'s")
        else:
            out.append(token.title())
    return " ".join(out)

def _absence_label(start_date, last_absent_date, all_day):
    if not start_date:
        return "out of office"

    next_week_start = today + timedelta(days=(7 - today.weekday()))
    if start_date <= today <= last_absent_date:
        label = "off next week" if today.weekday() >= 5 else "off today"
    elif start_date == tomorrow:
        label = "off next week" if today.weekday() >= 4 else "off tomorrow"
    elif start_date >= next_week_start:
        label = "off next week"
    else:
        label = "off " + _fmt_absence_date(start_date)

    return_date = next_workday(last_absent_date) if all_day else None
    if return_date and return_date > today:
        label += ", returns " + _fmt_absence_date(return_date)
    return label

# Defense-in-depth, not a substitute for the organizer-placeholder check
# above: after cleaning, reject anything that still reads as a department/
# system name rather than a person. Caught live, 10 Aug 2026 -- a real
# production run produced "People Department - Hr Systems" as a bogus
# absent "person" despite the organizer-placeholder pre-check, and a
# same-session, same-logic replication immediately after could NOT
# reproduce it (most likely an Outlook COM quirk specific to expanding a
# recurring series via IncludeRecurrences, not a pinned-down logic bug) --
# rather than keep chasing a non-reproducible trigger, this output-side
# check guards against the whole class of "organizer/name resolved to
# something obviously not a person" regardless of which mechanism causes it.
_NON_PERSON_NAME_TERMS = ["department", "systems", " team", "hr systems"]

def _looks_like_a_person(name):
    lower = name.lower()
    return not any(term in lower for term in _NON_PERSON_NAME_TERMS)

# Root-caused live, 21 Aug 2026 (see HANDOVER.md and
# begb0037admin/drew/memory/wi-absences-dedup-diagnosis-21aug.md): the
# old _add_absence() was unconditional first-write-wins, so when a
# person had more than one real, eligible calendar entry in the window
# the EARLIEST-processed one silently won and any later, more relevant
# one was dropped with no record of it anywhere. Replaced with a
# two-pass design: collect every eligible entry per person first, then
# resolve each person's full candidate list in one place.
def _resolve_person_name(organizer, sender_name, subject):
    # Preference order: (1) a real (non-placeholder) Organizer; (2) a
    # real (non-placeholder) PR_SENDER_NAME, captured at Phase 1 -- see
    # the comment above _get_pr_sender_name(); (3) the subject text, as
    # a last resort for entries with no distinguishing submitter at all
    # (e.g. bare "Kevin - A/L" bookings where even PR_SENDER_NAME is
    # just the placeholder department name).
    def _is_placeholder(n):
        n = n.lower()
        return "people department" in n or "hr systems" in n

    organizer = (organizer or "").strip()
    if len(organizer) >= 3 and not _is_placeholder(organizer):
        return organizer
    sender_name = (sender_name or "").strip()
    if len(sender_name) >= 3 and not _is_placeholder(sender_name):
        return sender_name
    return subject

absence_candidates = {}  # key -> list of {name, start, end, all_day, is_recurring}

def _collect_absence_candidate(name, start_date, last_absent_date, all_day, is_recurring):
    name = _clean_absence_name(name)
    key = _absence_key(name)
    if not key or len(name) < 3:
        return
    if not _looks_like_a_person(name):
        return
    absence_candidates.setdefault(key, []).append({
        "name": name, "start": start_date, "end": last_absent_date,
        "all_day": all_day, "is_recurring": is_recurring,
    })

week_absence_end = today + timedelta(days=8)
for item in calendar:
    subj = item.get("subject") or ""
    subj_lower = subj.lower()
    # Bare "AL" (no slash) counts as a leave keyword here too -- see the
    # _has_bare_al() comment above _DAY_VIEW_EXCLUDE_KEYWORDS for why this
    # needs standalone-word regex matching rather than a plain substring.
    # Fixes "Michael - AL" (and any future same-pattern entry) being
    # invisible in the sidebar Absences panel, not just the day-view.
    if not (any(kw in subj_lower for kw in ABSENCE_KEYWORDS) or _has_bare_al(subj)):
        continue
    if "absence reporting" in subj_lower or "sickness absence report" in subj_lower:
        continue

    start_date = _parse_iso_date(item.get("start"))
    end_date = _parse_iso_date(item.get("end")) or start_date
    if not start_date:
        continue

    all_day = bool(item.get("all_day"))
    last_absent_date = (end_date - timedelta(days=1)) if all_day and end_date > start_date else end_date
    if last_absent_date < today or start_date > week_absence_end:
        continue

    name_source = _resolve_person_name(item.get("organizer"), item.get("sender_name"), subj)
    _collect_absence_candidate(
        name_source, start_date, last_absent_date, all_day,
        bool(item.get("is_recurring"))
    )

# Resolve each person's candidate list into a single display entry.
#
# Kevin's explicit policy, 21 Aug 2026: a recurring "non-working
# day"/pattern match is a background-schedule signal, not a real
# absence, so it should only surface when NO real (non-recurring) entry
# exists for that person in the window -- real entries always take
# priority over recurring ones, not merely tie-break against them.
#
# Within whichever tier is used, entries that are genuinely continuous
# (no gap at all between one's last day and the next's first day) are
# merged into a single combined window -- true calendar-day adjacency
# only (next window's start is the same day as, or the very next day
# after, the previous window's last day); this was already correct in
# the first pass and is unchanged here. Entries that are NOT touching
# are treated as genuinely separate absence periods and are NEVER
# bridged into a fabricated combined span.
#
# SECOND PASS, 21 Aug 2026 -- corrects the first pass's selection rule.
# The first pass picked whichever window had the LATEST start date,
# reasoning (wrongly) that "most recent booking" was most relevant.
# Live re-verification the same day found real people with more than
# two genuinely separate windows in the eligible range (Kevin: three
# separate single-day A/L entries, Mon 24 / Wed 26 / Fri 28 Aug, each
# its own calendar entry, confirmed genuinely separate by Kevin
# directly -- not a data error, not one span), and "latest start wins"
# fabricates a bridge past real gaps to whichever window happens to
# start last, e.g. producing "off next week, returns Thursday 27
# August" for Kevin by silently discarding the nearer, actually-current
# Monday 24 window. Corrected rule: prefer whichever window genuinely
# covers TODAY, if any; otherwise the window with the EARLIEST start
# date that is still upcoming (the soonest-relevant one) -- never the
# latest. Re-verified live against fresh Outlook COM data, 21 Aug 2026:
# this is the only rule that reproduces Kevin's own confirmed real
# dates for his own three entries (today, Fri 21 Aug, falls in none of
# them, so the soonest -- Mon 24, returning Tue 25 -- is what's
# surfaced) and is applied uniformly for every person, not special-
# cased to Kevin's or anyone else's name -- so it also corrects the
# same "latest wins" bug for Simon Burford and Anthony Kong, who were
# each independently found live to now have a second, separate real
# window the first pass would have wrongly bridged to.
#
# Any non-surfaced real window is not silently discarded -- it's
# written to the run log below, since the current one-line-per-person
# absences data shape can only display one line per person (a real
# limitation of today's data model, not addressed by this fix -- see
# HANDOVER.md).
def _gap_is_all_weekend(prev_end, next_start):
    # True only if EVERY day strictly between prev_end and next_start is a
    # Saturday or Sunday. No UK bank-holiday list or other non-working-day
    # concept exists anywhere in this codebase to extend this with --
    # checked live, 21 Aug 2026 (fetch_inbox.py's own next_workday() only
    # ever skips weekday() >= 5, nothing else; the "non-working day" subject
    # keyword elsewhere in this file is a different, personal-pattern
    # concept -- see the tier-priority comment above -- not a public
    # holiday calendar). Weekend-only is therefore the correct scope, not a
    # cut corner. A zero-day gap (adjacent/overlapping) trivially returns
    # True since the loop body never runs.
    d = prev_end + timedelta(days=1)
    while d < next_start:
        if d.weekday() < 5:
            return False
        d += timedelta(days=1)
    return True

# THIRD PASS, 21 Aug 2026 -- refines the second pass's adjacency rule.
# The second pass merged two real windows only on exact calendar-day
# adjacency (zero gap). Kevin directly confirmed that's too strict: Michael
# O'Sullivan has a real Fri 21 Aug entry and a real Mon 24 Aug entry with a
# two-day numeric gap (Sat 22, Sun 23) -- but from his actual perspective
# that's ONE continuous absence, since the gap is entirely non-working days
# anyway. The second pass's strict-adjacency rule wrongly kept these
# separate and picked "off today, returns Monday 24 August" (the current
# Friday window, discarding the nearer relevance of Monday). Kevin's own
# words, relayed via the coordinator: bridge two real entries only if
# EVERY day in the gap is a non-working day (weekend, or an existing
# non-working-day concept if the codebase had one -- it doesn't, see
# _gap_is_all_weekend() above). If the gap contains any real working day,
# keep the entries separate, per the second pass's already-correct logic.
#
# Verified live against real 2026 calendar dates (see
# begb0037admin/drew/memory/wi-absences-dedup-third-pass-21aug.md for the
# full day-of-week table): Michael's gap (Sat 22 / Sun 23) is 100% weekend
# -> now bridges into one Fri 21-Mon 24 window, correctly producing "off
# today, returns Tuesday 25 August". Kevin's gaps are NOT pure weekends --
# Mon 24 -> Wed 26 has Tue 25 (a Tuesday) in the gap, and Wed 26 -> Fri 28
# has Thu 27 (a Thursday) in the gap -- so his three entries correctly stay
# separate and his label is unaffected.
def _merge_adjacent_windows(cands):
    windows = sorted((dict(c) for c in cands), key=lambda c: c["start"])
    merged = []
    for c in windows:
        if merged and (
            c["start"] <= merged[-1]["end"] + timedelta(days=1)
            or _gap_is_all_weekend(merged[-1]["end"], c["start"])
        ):
            merged[-1]["end"] = max(merged[-1]["end"], c["end"])
            merged[-1]["all_day"] = merged[-1]["all_day"] or c["all_day"]
        else:
            merged.append(c)
    return merged

for key, cands in absence_candidates.items():
    real = [c for c in cands if not c["is_recurring"]]
    recurring = [c for c in cands if c["is_recurring"]]
    tier = real if real else recurring
    if not tier:
        continue
    windows = _merge_adjacent_windows(tier)
    current_windows = [w for w in windows if w["start"] <= today <= w["end"]]
    chosen = current_windows[0] if current_windows else min(windows, key=lambda w: w["start"])
    if len(windows) > 1:
        for w in windows:
            if w is chosen:
                continue
            log(f"Phase absences - {chosen['name']}: {len(windows)} separate "
                f"{'real' if real else 'recurring'} windows in window; "
                f"surfacing {chosen['start']}..{chosen['end']} "
                f"({'covers today' if current_windows else 'soonest upcoming'}), "
                f"not dropping {w['start']}..{w['end']} (see HANDOVER.md -- "
                f"one-line-per-person display can't show both)")
    label = _absence_label(chosen["start"], chosen["end"], chosen["all_day"])
    absence_map[key] = chosen["name"] + " - " + label if label else chosen["name"]

# No email-OOO fallback -- calendar-only sourcing per Kevin's explicit
# decision, 10 Aug 2026 (see comment above ABSENCE_KEYWORDS). Every entry
# below now has a real calendar-verified date; "date unknown" can no longer
# appear in this list at all.
absences = sorted(absence_map.values())

# Priority actions -- pulled from Command Centre tasks.json.
# COMMAND_CENTRE_REPO / COMMAND_CENTRE_PATH / AUTO_PROMOTE_NEW_TASKS and the
# loader itself are hoisted to module scope (near _ai_create) so the early
# combined claude_code call can use them; _cc_load_priorities() is idempotent,
# so for the api path this is the first + only call.
_cc_load_priorities()

# Phase 3.5 - AI triage: which emails should become Command Centre tasks
log("Phase 3.5 - triaging inbox for task suggestions...")
suggestions = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "new_tasks":    [],
    "task_updates": []
}
# Dedupe ledger - emails already applied to Command Centre tasks
ledger = {"applied": {}, "promoted": {}}
if GITHUB_PAT:
    try:
        _lurl = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/triage_ledger.json"
        _lreq = urllib.request.Request(_lurl, headers={"Authorization": f"token {GITHUB_PAT}", "User-Agent": "work-inbox-script"})
        with urllib.request.urlopen(_lreq) as r:
            ledger = json.loads(base64.b64decode(json.loads(r.read())["content"]).decode("utf-8"))
        if "applied" not in ledger:
            ledger["applied"] = {}
        if "promoted" not in ledger:
            ledger["promoted"] = {}
    except Exception:
        pass
try:
    if not anthropic_available:
        raise RuntimeError("skipped because Anthropic is unavailable")
    task_summaries = []
    task_list = cc_content if isinstance(cc_content, list) else cc_content.get("tasks", [])
    for t in task_list:
        task_summaries.append({
            "id":          t.get("id", ""),
            "title":       t.get("title", ""),
            "description": (t.get("description") or "")[:300],
            "emailRef":    t.get("emailRef", "")
        })
    if not task_summaries:
        raise Exception("Command Centre tasks unavailable - skipping triage")

    email_candidates = []
    for m in inbox:
        if categorise(m) in ("urgent", "needs"):
            email_candidates.append({
                "subject":      m.get("subject", ""),
                "from":         m.get("from", ""),
                "received":     (m.get("received", "") or "")[:16],
                "body_preview": re.sub(r"<\?\s*https?://\S+>?", "[link]", (m.get("body_preview") or ""))[:150],
                "entry_id":     m.get("entry_id", "")
            })

    for s in sent[:30]:
        email_candidates.append({
            "subject":      s.get("subject", ""),
            "from":         "Kevin (sent to: " + (s.get("to") or "") + ")",
            "received":     (s.get("sent", "") or "")[:16],
            "body_preview": re.sub(r"<\?\s*https?://\S+>?", "[link]", (s.get("body_preview") or ""))[:150],
            "entry_id":     s.get("entry_id", ""),
            "direction":    "sent"
        })

    api_emails = [{"n": i, "direction": e.get("direction", "received"),
                   "subject": e["subject"], "from": e["from"],
                   "received": e["received"], "body_preview": e["body_preview"]}
                  for i, e in enumerate(email_candidates)]

    TRIAGE_SYSTEM = _SYS_TRIAGE  # verbatim; hoisted to module scope

    triage_user = (
        f"Today is {today_str}. Tomorrow (next working day) is {tomorrow_str}.\n\n"
        f"EXISTING TASKS:\n{json.dumps(task_summaries, indent=1, ensure_ascii=True)}\n\n"
        f"EMAILS (received urgent/needs + sent by Kevin, last 7 days):\n{json.dumps(api_emails, indent=1, ensure_ascii=True)}"
    )

    t_resp = _ai_create(
        model      = "claude-haiku-4-5",
        max_tokens = 8000,
        system     = TRIAGE_SYSTEM,
        messages   = [{"role": "user", "content": triage_user}],
        _phase     = "task_triage",
    )
    t_raw = t_resp.content[0].text.strip()
    if t_raw.startswith("```"):
        t_raw = "\n".join(t_raw.split("\n")[1:])
    if t_raw.endswith("```"):
        t_raw = "\n".join(t_raw.split("\n")[:-1])
    t_out = json.loads(t_raw)

    task_by_id = {t["id"]: t for t in task_summaries}
    suppressed_no_action = 0
    for nt in t_out.get("new_tasks", [])[:12]:
        i = nt.get("email_n")
        if not isinstance(i, int) or not (0 <= i < len(email_candidates)):
            continue
        src = email_candidates[i]
        # 12 Aug 2026: don't let Phase 3.5's own (separate) AI triage call
        # spawn a brand-new Command Centre task suggestion from an email
        # this same run's Phase 3.2/3.3 already AI-confirmed is genuinely
        # nothing for Kevin to act on. Reuses `_noise_demoted_entry_ids`
        # (built by Phase 3.3/3.3b above) rather than re-running a second
        # classification pass. Deliberately does NOT filter task_updates
        # below -- a no_action_needed email can still be genuine, useful
        # progress information against an EXISTING already-tracked task
        # (e.g. "vendor confirms shipment date") even though Kevin
        # personally has nothing to do about it, so it stays a valid
        # candidate for that purpose; only brand-new task proposals are
        # noise here, matching the same demotion logic applied elsewhere.
        if src.get("entry_id") and src["entry_id"] in _noise_demoted_entry_ids:
            suppressed_no_action += 1
            continue
        suggestions["new_tasks"].append({
            "title":         nt.get("title", ""),
            "tier":          nt.get("tier") if nt.get("tier") in ("today", "tomorrow", "week") else "week",
            "description":   nt.get("description", ""),
            "email_subject": src["subject"],
            "email_from":    src["from"],
            "received":      src["received"],
            "entry_id":      src["entry_id"]
        })
    for tu in t_out.get("task_updates", [])[:20]:
        i   = tu.get("email_n")
        tid = tu.get("task_id", "")
        if not isinstance(i, int) or not (0 <= i < len(email_candidates)) or tid not in task_by_id:
            continue
        if email_candidates[i]["entry_id"] + "_" + tid in ledger.get("applied", {}):
            continue
        src = email_candidates[i]
        suggestions["task_updates"].append({
            "task_id":       tid,
            "task_title":    task_by_id[tid]["title"],
            "note":          tu.get("note", ""),
            "email_subject": src["subject"],
            "email_from":    src["from"],
            "received":      src["received"],
            "entry_id":      src["entry_id"]
        })
    print(f"Phase 3.5 done - new:{len(suggestions['new_tasks'])} (suppressed_no_action:{suppressed_no_action}) updates:{len(suggestions['task_updates'])}")
except Exception as e:
    if anthropic_available:
        print(f"WARNING: Phase 3.5 triage failed - {e}")
    else:
        print("Phase 3.5 skipped - Anthropic is unavailable")


# -- Assemble final briefing --

# Desktop-toast helper for internal phase failures this script otherwise
# swallows by design (graceful degradation -- see the many bare "except
# Exception" blocks throughout this file: a downstream phase failing must
# never take down the primary briefing deliverable). That design is correct
# for the pipeline's own resilience, but it means a real failure here (e.g.
# Phase 3.6's Command Centre sync) currently produces nothing Kevin will
# actually see -- fetch_inbox.py still exits 0, so the existing end-of-run
# Show-TaskNotification.ps1 call in "Run Inbox Briefing Hidden.vbs" reports
# "Success" regardless. This reuses that exact same notification mechanism
# (Show-TaskNotification.ps1 / BurntToast) mid-run instead of inventing a
# new one, for phases specifically flagged as needing real visibility (see
# HANDOVER.md, Phase 2 build 20 Aug 2026).
#
# Writes a small dedicated one-line detail file rather than pointing
# Show-TaskNotification.ps1 at the shared inbox_briefing_last_run.log --
# that log keeps growing for the rest of this run, and Get-LogTailDetail's
# regex ("Error|Traceback|Exception|com_error|Call was rejected") is not
# guaranteed to match every exception's str() text, so relying on it against
# a shared, still-growing log risks surfacing the wrong line by the time the
# toast script actually reads it. A dedicated one-line file makes the detail
# text deterministic regardless of what runs afterward.
#
# Best-effort only, by design: a failure to raise the toast must never mask
# or replace the original exception already being handled by the caller's
# own except block, and must never itself crash the run.
#
# _notify_phase_failure() and NOTIFY_SCRIPT_PATH were MOVED to the top of
# this file on 2026-08-28 (Drew) so connect_to_outlook() can also use them.
# The definition now lives just after log(); this is only a pointer.

# Phase 3.6 - apply task updates directly to Command Centre tasks.json
def _gh_get(url, headers):
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
        return json.loads(r.read())

def _gh_put(url, headers, message, content_bytes, sha=None):
    payload = {"message": message,
               "content": base64.b64encode(content_bytes).decode("ascii")}
    if sha:
        payload["sha"] = sha
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"),
                                 headers=headers, method="PUT")
    with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
        return json.loads(r.read())

def _backup_briefing_before_write(remote_meta, headers):
    if not remote_meta or not remote_meta.get("content"):
        return
    backup_bytes = base64.b64decode(remote_meta["content"])
    backup_path = f"data/archive/briefing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    backup_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{backup_path}"
    _gh_put(
        backup_url,
        headers,
        f"backup: briefing before refresh {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        backup_bytes
    )
    print(f"Phase 4 backup created - {backup_path}")

if AI_PARALLEL and (suggestions["task_updates"] or suggestions["new_tasks"]):
    log("Phase 3.6 - PARALLEL VALIDATION MODE: NOT applying task updates to Command Centre / NOT writing triage_ledger.json (comparison only).")
if PUSH_ENABLED and (suggestions["task_updates"] or suggestions["new_tasks"]):
    try:
        gh_headers = {"Authorization": f"token {GITHUB_PAT}",
                      "Content-Type":  "application/json",
                      "User-Agent":    "work-inbox-script"}
        cc_tasks_url = f"https://api.github.com/repos/{COMMAND_CENTRE_REPO}/contents/{COMMAND_CENTRE_PATH}"
        cc_meta   = _gh_get(cc_tasks_url, gh_headers)
        tasks_doc = json.loads(base64.b64decode(cc_meta["content"]).decode("utf-8"))

        # Mandatory daily backup before any write to tasks.json
        backup_path = f"Archive/tasks_backup_{datetime.now().strftime('%Y%m%d')}.json"
        backup_url  = f"https://api.github.com/repos/{COMMAND_CENTRE_REPO}/contents/{backup_path}"
        try:
            _gh_get(backup_url, gh_headers)
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
            _gh_put(backup_url, gh_headers,
                    f"backup: tasks.json {datetime.now().strftime('%Y-%m-%d')}",
                    base64.b64decode(cc_meta["content"]))
            print(f"Phase 3.6 - daily backup created: {backup_path}")

        stamp   = datetime.now().strftime("%d %b %Y")
        applied = 0
        task_list = tasks_doc if isinstance(tasks_doc, list) else tasks_doc.get("tasks", [])
        for task in task_list:
            for upd in suggestions["task_updates"]:
                if task.get("id") == upd["task_id"]:
                    action_text = f"[{stamp}] {upd['note']} (email: {upd['email_from']} - {upd['email_subject']})"
                    actions = task.setdefault("actions", [])
                    if action_text in actions:
                        break
                    actions.append(action_text)
                    task["entryId"] = upd["entry_id"]
                    applied += 1
                    break

        # Auto-promote new task suggestions straight into tasks.json.
        # Guarded four ways so a task is never created twice: the promoted
        # ledger, an existing task already carrying that entryId, an exact
        # case-insensitive title match, and (added 12 Aug 2026) a fuzzy
        # near-duplicate title match against every existing task title.
        # The fuzzy guard exists because the exact-match guard alone missed
        # a confirmed live duplicate pair where two separate emails about
        # the same request produced two near-identical titles ("Advise on
        # GLAM joining 38-day balance departments scheme" vs "Advise Marie
        # on GLAM joining 38-day balance scheme", SequenceMatcher ratio
        # 0.83). Threshold 0.8 chosen empirically against live task data:
        # it catches every confirmed live duplicate pair (0.83-1.0) while
        # staying clear of the closest known false positive -- two
        # genuinely different meetings that share generic words (0.68).
        FUZZY_DUP_THRESHOLD = 0.8
        def _fuzzy_duplicate_of(title, other_titles):
            for other in other_titles:
                if SequenceMatcher(None, title.lower(), other.lower()).ratio() >= FUZZY_DUP_THRESHOLD:
                    return other
            return None

        existing_entry_ids  = {t.get("entryId") for t in task_list if t.get("entryId")}
        existing_titles     = {(t.get("title") or "").strip().lower() for t in task_list}
        existing_title_list = [(t.get("title") or "").strip() for t in task_list if t.get("title")]
        promoted = 0
        fuzzy_skipped = 0
        for nt in (suggestions["new_tasks"] if AUTO_PROMOTE_NEW_TASKS else []):
            eid = nt.get("entry_id", "")
            if not eid or eid in ledger.get("promoted", {}) or eid in existing_entry_ids:
                continue
            title = (nt.get("title") or "").strip()
            if title.lower() in existing_titles:
                continue
            dup_of = _fuzzy_duplicate_of(title, existing_title_list)
            if dup_of:
                fuzzy_skipped += 1
                print(f"Phase 3.6 - skipped near-duplicate task suggestion: '{title}' looks like existing '{dup_of}'")
                continue
            new_id = "t" + datetime.now().strftime("%y%m%d%H%M%S") + str(promoted)
            task_list.append({
                "id":          new_id,
                "title":       nt["title"],
                "tier":        nt["tier"],
                "source":      f"Inbox - {nt['email_from']}, {nt['received']}",
                "emailRef":    nt.get("email_subject", ""),
                "entryId":     eid,
                "summary":     "",
                "description": nt.get("description", ""),
                "origin":      "inbox-auto",
                "actions":     [f"[{stamp}] Auto-created from inbox triage (email: {nt['email_from']} - {nt.get('email_subject','')})."]
            })
            existing_entry_ids.add(eid)
            existing_titles.add(title.lower())
            existing_title_list.append(title)
            nt["auto_promoted"] = True
            promoted += 1

        if applied or promoted:
            bits = []
            if applied:
                bits.append(f"apply {applied} task update(s)")
            if promoted:
                bits.append(f"add {promoted} new task(s)")
            _gh_put(cc_tasks_url, gh_headers,
                    f"inbox: {' + '.join(bits)} {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                    json.dumps(tasks_doc, indent=2, ensure_ascii=False).encode("utf-8"),
                    cc_meta["sha"])
            print(f"Phase 3.6 done - {applied} update(s), {promoted} new task(s) applied to Command Centre" + (f" ({fuzzy_skipped} near-duplicate suggestion(s) skipped)" if fuzzy_skipped else ""))
            for u in suggestions["task_updates"]:
                ledger["applied"][u["entry_id"] + "_" + u["task_id"]] = datetime.now().strftime("%Y-%m-%d")
            for nt in suggestions["new_tasks"]:
                if nt.get("auto_promoted"):
                    ledger["promoted"][nt["entry_id"]] = datetime.now().strftime("%Y-%m-%d")
            ledger_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/triage_ledger.json"
            l_sha = None
            try:
                l_sha = _gh_get(ledger_url, gh_headers).get("sha")
            except urllib.error.HTTPError as e:
                if e.code != 404:
                    raise
            _gh_put(ledger_url, gh_headers, "chore: update triage ledger",
                    json.dumps(ledger, indent=1).encode("utf-8"), l_sha)
        suggestions["applied_updates"] = suggestions.pop("task_updates")
    except Exception as e:
        print(f"WARNING: Phase 3.6 apply failed - {e}")
        _notify_phase_failure("Work Inbox Briefing - Command Centre sync", f"{type(e).__name__}: {e}")


# Phase 3.7 - AI summaries for priority tasks
log("Phase 3.7 - generating AI task summaries...")
all_priorities = priorities_today + priorities_tomorrow + priorities_week
if all_priorities and anthropic_available:
    try:
        tasks_for_summary = [
            {
                "id":          e["id"],
                "title":       e["text"],
                "description": (e.get("description") or "")[:300],
                "actions":     e.get("actions", [])[-5:]
            }
            for e in all_priorities if e.get("id")
        ]
        SUMMARY_SYSTEM = _SYS_TASK_SUMMARY  # verbatim; hoisted to module scope
        summary_user = (
            f"Today is {today_str}.\n\n"
            f"TASKS:\n{json.dumps(tasks_for_summary, indent=1, ensure_ascii=True)}"
        )
        s_resp = _ai_create(
            model      = "claude-haiku-4-5",
            max_tokens = 4096,
            system     = SUMMARY_SYSTEM,
            messages   = [{"role": "user", "content": summary_user}],
            _phase     = "task_summary",
        )
        s_raw = s_resp.content[0].text.strip()
        if s_raw.startswith("```"):
            s_raw = "\n".join(s_raw.split("\n")[1:])
        if s_raw.endswith("```"):
            s_raw = "\n".join(s_raw.split("\n")[:-1])
        summaries = json.loads(s_raw)
        for entry in all_priorities:
            tid = entry.get("id", "")
            if tid in summaries:
                entry["ai_summary"] = summaries[tid]
        print(f"Phase 3.7 done - {len(summaries)} summaries generated")
    except Exception as e:
        print(f"WARNING: Phase 3.7 AI summaries failed - {e}")
elif all_priorities:
    print("Phase 3.7 skipped - Anthropic is unavailable")


# Pre-build cal items so Phase 3.8 can annotate them before briefing dict is assembled
cal_today_items    = build_cal_items(cal_today)
cal_tomorrow_items = build_cal_items(cal_tomorrow)
cal_day2_items     = build_cal_items(cal_day2)
cal_day3_items     = build_cal_items(cal_day3)

# Attach a Command Centre task id to a calendar item, when one genuinely
# exists, so the dashboard's "CC ->" link on that meeting can deep-link
# straight to the matching task (command-centre's js/app.js already reads
# window.location.hash and highlights/scrolls to '#card-'+hash -- confirmed
# by reading that code, not assumed) instead of just landing on the CC
# homepage with nothing highlighted. Kevin's explicit ask, 10 Aug 2026:
# "it should high[light] the item so i can drill dowwn into the email if
# required - one links to the other."
#
# Deliberately conservative: only an EXACT (case-insensitive) match against
# a not-done task's emailRef counts -- confirmed live against real
# tasks.json that several tasks carry the verbatim meeting title in
# emailRef (e.g. "Sickness Absence Survey working group", "Confidential -
# OH Consultation", "Oxford University Evo Pre project meeting"). task.source
# also often names a meeting but with a trailing "DD/MM" date and no way to
# tell which week's occurrence of a *recurring* meeting it refers to (e.g.
# "HR Systems Managers Meeting 24/06" could easily be a stale prior
# occurrence of a meeting that recurs weekly) -- deliberately NOT matched
# against, to avoid deep-linking a real person to the wrong week's task.
# If more than one not-done task shares the identical emailRef, that's
# ambiguous and no link is attached rather than guessing.
_cc_task_list_for_matching = cc_content if isinstance(cc_content, list) else cc_content.get("tasks", [])

def _match_cc_task_id(meeting_title):
    title_lower = (meeting_title or "").strip().lower()
    if not title_lower:
        return None
    matches = []
    for t in _cc_task_list_for_matching:
        if t.get("done"):
            continue
        email_ref = (t.get("emailRef") or "").strip().lower()
        if email_ref and email_ref == title_lower:
            tid = t.get("id")
            if tid and tid not in matches:
                matches.append(tid)
    return matches[0] if len(matches) == 1 else None

def _attach_cc_task_ids(items):
    for it in items:
        tid = _match_cc_task_id(it.get("title"))
        if tid:
            it["ccTaskId"] = tid
    return items

_attach_cc_task_ids(cal_today_items)
_attach_cc_task_ids(cal_tomorrow_items)
_attach_cc_task_ids(cal_day2_items)
_attach_cc_task_ids(cal_day3_items)

# -- Phase 3.7b -- Fetch recent Granola meeting notes for calendar context --
# GRANOLA_API_KEY, _granola_context, _granola_keywords, _granola_fetch,
# _non_all_day_candidates and the fetch loop (_cc_fetch_granola) are hoisted
# to module scope so the early combined claude_code call can use them.
#
# The idx-fixed candidate list carries "idx" (sequential 0-based within its
# day, the ONLY index shown to the model) and "real_idx" (true position in
# cal_<day>_items, used only for Phase 3.8's local write-back) -- fixes the
# calendar-summary offset bug root-caused 4 Aug 2026
# (memory/calendar-summary-offset-bug.md).
if AI_BACKEND != "claude_code":
    _all_day_candidates = (
        _non_all_day_candidates(cal_today_items, "today") +
        _non_all_day_candidates(cal_tomorrow_items, "tomorrow") +
        _non_all_day_candidates(cal_day2_items, "day2") +
        _non_all_day_candidates(cal_day3_items, "day3")
    )
    _cc_fetch_granola(_all_day_candidates)
# else: _all_day_candidates + _granola_context were built by the early
# _cc_build_cal_candidates_early() / _cc_fetch_granola() before the combined call.

# -- Phase 3.8 -- AI prep summaries for all 4 calendar day-view columns --
# Reuse the same idx-fixed candidate list -- "idx" (sequential, model-facing)
# is what gets sent to and read back from the AI; "real_idx" (true position
# in cal_X_items) is only used for the write-back below, never sent to the
# model. Fixes the calendar-summary offset bug -- see comment above
# _non_all_day_candidates().
_cal_for_summary = [
    dict(c, prev_meeting_notes=_granola_context.get(f"{c['day']}_{c['idx']}", {}).get("summary", ""))
    for c in _all_day_candidates
]
if _cal_for_summary and anthropic_available:
    try:
        CAL_SUM_SYSTEM = _SYS_CAL  # verbatim; hoisted to module scope
        _cal_user = (
            f"Today is {today_str}.\n\n"
            f"MEETINGS:\n{json.dumps(_cal_for_summary, indent=1, ensure_ascii=True)}"
        )
        _cs_resp = _ai_create(
            model      = "claude-haiku-4-5",
            max_tokens = 900,
            system     = CAL_SUM_SYSTEM,
            messages   = [{"role": "user", "content": _cal_user}],
            _phase     = "calendar_prep",
        )
        _cs_raw = _cs_resp.content[0].text.strip()
        if _cs_raw.startswith("```"): _cs_raw = "\n".join(_cs_raw.split("\n")[1:])
        if _cs_raw.endswith("```"):   _cs_raw = "\n".join(_cs_raw.split("\n")[:-1])
        _cs_map = json.loads(_cs_raw)
        _CAL_DAY_TARGETS = {
            "today": cal_today_items, "tomorrow": cal_tomorrow_items,
            "day2": cal_day2_items, "day3": cal_day3_items,
        }
        for item in _cal_for_summary:
            key = f"{item['day']}_{item['idx']}"
            if key in _cs_map:
                target = _CAL_DAY_TARGETS[item["day"]]
                target[item["real_idx"]]["summary"] = _cs_map[key]
        print(f"Phase 3.8 done - {len(_cs_map)} calendar summaries generated")
    except Exception as e:
        print(f"WARNING: Phase 3.8 calendar summaries failed - {e}")
elif _cal_for_summary:
    print("Phase 3.8 skipped - Anthropic is unavailable")

if same_briefing_date(existing_briefing, today_str):
    _preserved_cal = (
        preserve_existing_calendar_summaries(existing_briefing, "calToday", cal_today_items) +
        preserve_existing_calendar_summaries(existing_briefing, "calTomorrow", cal_tomorrow_items) +
        preserve_existing_calendar_summaries(existing_briefing, "calDay2", cal_day2_items) +
        preserve_existing_calendar_summaries(existing_briefing, "calDay3", cal_day3_items)
    )
    if _preserved_cal:
        print(f"Phase 3.8 preservation - reused {_preserved_cal} existing same-day calendar summaries")

# Build calFull -- Mon through Fri of the current working week
def _week_workdays(ref):
    mon = ref - timedelta(days=ref.weekday())
    return [mon + timedelta(days=i) for i in range(5)]

calFull = []
for _wd in _week_workdays(today):
    _day_items = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == _wd]
    calFull.append({
        "date":    _wd.strftime("%Y-%m-%d"),
        "label":   _wd.strftime("%A") + " " + str(_wd.day) + " " + _wd.strftime("%b"),
        "items":   build_cal_items(_day_items),
        "isToday": _wd == today
    })

if same_briefing_date(existing_briefing, today_str) and not absences and existing_briefing.get("absences"):
    absences = sorted(existing_briefing.get("absences", []))
    print(f"Absence preservation - reused {len(absences)} existing same-day absence(s)")

# -- Phase 3.9 -- Needs/Urgent scroll-out persistence, v2 (20 Aug 2026) --
# REBUILT from a fresh design pass, not a cherry-pick of the old revert. The
# original version of this (17 Aug 2026, commits 5216d9fd/640b44ee) was
# fully reverted the same night at Kevin's explicit request -- but the
# revert was about resetting a messy same-night investigation, not a flaw
# found in this mechanism itself: by the time of the revert, Phase 3.9 had
# already been proven live end-to-end (a real "carried:2" production run,
# plus a real production round-trip proving the done-tick resolution signal
# worked). Kevin separately asked for a fresh design pass rather than a
# quick re-patch, so the core mechanism below was deliberately re-evaluated,
# not just restored -- kept where it was sound, changed in two places (see
# "What changed from v1" below).
#
# Problem (unchanged from v1, still real, still verified): Phase 1 only
# ever pulls the 50 newest Outlook items. An item the AI correctly triages
# into Urgent/Needs has no durable state once it scrolls out of that
# window -- it just silently vanishes from every tier, even though nothing
# about it was ever resolved. Real-world case that exposed this: an Alan
# Quirke/Access Group vendor email, correctly triaged Needs on 31 Jul,
# gone without a trace by 2 Aug (see wi-quirke-needs-tier-scrollout-17aug.md
# in the `drew` repo for the full archived-briefing evidence trail).
#
# Mechanism:
# 1. Every run, after Phase 3.2/3.3/3.3b have finished (so cards already
#    carry their AI summary/needs_reply verdict), snapshot every currently-
#    live Urgent/Needs card into triage_ledger.json under
#    "tracked_needs_urgent", keyed by entry_id.
# 2. Any previously-tracked entry_id NOT in this run's fresh urgent/needs
#    (i.e. it scrolled out of the pull) is checked against three
#    independent resolution signals, in this order:
#      a. Outlook GetItemFromID -> item.Parent.EntryID vs the Inbox's own
#         EntryID. Moved out of the Inbox (filed/archived/deleted) ->
#         Kevin has dealt with it -> resolved, drop. Lookup throwing for
#         any reason -> UNKNOWN, fail OPEN (still carried) -- confirmed
#         live before v1 was built that Outlook's LastVerbExecuted
#         property is not reliably readable through this COM binding, so
#         it is deliberately not used; failing open on any lookup
#         uncertainty is the safer failure mode given the bug this fixes
#         is silent vanishing, not over-carrying.
#      b. Command Centre tasks.json: a task whose entryId matches and is
#         done:true -> resolved via the other surface, drop.
#      c. The dashboard's own data/ticks.json: a true-valued
#         'eid_<entry_id>' key -> Kevin ticked it done in the UI -> drop.
#         This signal did not exist in the very first same-day cut of v1
#         and was added hours later that same night after a live "mark
#         done, refresh, it comes back" incident -- it ships here from the
#         start this time, not as an emergency same-night patch (see "What
#         changed from v1", point 1).
# 3. Safety valve only, not the primary mechanism: an item carried for more
#    than 90 days without resolving gets a visible WARNING logged, but is
#    still carried -- Kevin was explicit these should not silently vanish,
#    so this does not reintroduce a silent time-based drop.
#
# What changed from v1:
# 1. Tick-key stability ships together with this fix, not as a same-night
#    follow-on patch. v1's real incident on 17 Aug was caused by exactly
#    this ordering problem: Phase 3.9 shipped first, keyed against
#    js/app.js's THEN-positional tick storage keys
#    (`<calendar-day>_pri_<sec>_<i>`), which broke the moment an item was
#    reordered, dragged, or carried across a day boundary by this very
#    mechanism -- carrying an item across days was new behaviour Phase 3.9
#    itself introduced, and it collided with a pre-existing tick-key bug
#    nothing had previously exercised hard enough to surface. This time,
#    js/app.js's card/tick identity was made day-independent and
#    position-independent (same _priGetKey()-derived id used for the DOM
#    element AND the tick storage key, see the drag-and-drop architecture
#    rework in the same changeset) BEFORE this mechanism goes live, so
#    resolution signal (c) above can never resurrect a ticked item on
#    reorder or day rollover.
# 2. A dry-run safety valve (WI_PHASE39_DRY_RUN env var) -- v1 had none;
#    the only way to validate it was to let it write to production
#    triage_ledger.json/briefing.json and inspect the result after the
#    fact. This version can run the full real Outlook/CC/tick resolution
#    logic against live data and log exactly what it WOULD do, with zero
#    writes, so a session (or Kevin) can verify correctness against real
#    live data before ever letting it touch main. Used to verify this
#    build itself, see the `drew` repo's memory entry for this session.
# 3. Carried-forward cards are tagged `_carried_forward: true` in the
#    injected copy -- pure metadata, not rendered by index.html/app.js
#    (confirmed: card rendering only reads specific known fields, unknown
#    extra fields are inert), added purely so a future debugging session
#    can immediately tell "was this card fresh this run or resurrected
#    from the ledger" without cross-referencing the ledger by hand. No new
#    UI tier, no visible badge, no AI-prompt change -- same narrow scoping
#    Kevin gave the original fix.
#
# Deliberately NOT rebuilt as part of this: the one-time 48-item historical
# backfill sweep v1 also did (re-scanning ~101 archived briefings and
# re-verdicting historical candidates with the live AI prompt). That was a
# one-off recovery action for the backlog that had already accumulated
# during the pre-fix gap, not part of the ongoing mechanism -- rebuilding
# it wasn't asked for in this pass and isn't needed for the mechanism
# itself to work correctly going forward; flagged to Kevin as a separate,
# optional follow-up if he wants that historical recovery re-run.
_WI_PHASE39_DRY_RUN = (os.environ.get("WI_PHASE39_DRY_RUN", "").strip().lower() in ("1", "true", "yes")) or AI_PARALLEL
try:
    _persist_ledger_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/triage_ledger.json"
    _persist_ro_headers = {"Authorization": f"token {GITHUB_PAT}", "User-Agent": "work-inbox-script"} if GITHUB_PAT else None
    _persist_ledger = {"applied": {}, "promoted": {}, "tracked_needs_urgent": {}}
    _persist_sha = None
    if GITHUB_PAT:
        try:
            _pmeta = _gh_get(_persist_ledger_url, _persist_ro_headers)
            _persist_sha = _pmeta.get("sha")
            _persist_ledger = json.loads(base64.b64decode(_pmeta["content"]).decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
        except Exception as e:
            print(f"WARNING: Phase 3.9 could not read triage_ledger.json, starting fresh - {e}")
    if "tracked_needs_urgent" not in _persist_ledger:
        _persist_ledger["tracked_needs_urgent"] = {}
    tracked = _persist_ledger["tracked_needs_urgent"]

    # Command Centre done-task entry_ids, for the resolution cross-check.
    _cc_done_entry_ids = set()
    try:
        _cc_task_list = cc_content if isinstance(cc_content, list) else cc_content.get("tasks", [])
        for _t in _cc_task_list:
            if _t.get("done") and _t.get("entryId"):
                _cc_done_entry_ids.add(_t["entryId"])
    except Exception:
        pass

    # Dashboard-ticked-done entry_ids, for the resolution cross-check. Reads
    # data/ticks.json directly (this script has no access to the dashboard's
    # own in-browser copy) and treats a true-valued 'eid_<entry_id>' key as
    # a genuine resolution, same as the CC-done check above.
    _ticked_done_entry_ids = set()
    try:
        _ticks_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/ticks.json"
        _ticks_meta = _gh_get(_ticks_url, _persist_ro_headers)
        _ticks_doc = json.loads(base64.b64decode(_ticks_meta["content"]).decode("utf-8"))
        for _k, _v in (_ticks_doc.get("ticks") or {}).items():
            if _v is True and isinstance(_k, str) and _k.startswith("eid_"):
                _ticked_done_entry_ids.add(_k[4:])
    except Exception as e:
        print(f"WARNING: Phase 3.9 could not read data/ticks.json for done-tick cross-check - {e}")

    today_iso = datetime.now().strftime("%Y-%m-%d")

    # 1. Checkpoint every live card this run (refreshes cached content and
    #    first_tracked/last_confirmed bookkeeping for anything still in the
    #    fresh pull).
    for _tier_name, _tier_list in (("urgent", urgent), ("needs", needs)):
        for c in _tier_list:
            eid = c.get("entry_id")
            if not eid:
                continue
            prev = tracked.get(eid, {})
            tracked[eid] = {
                "tier": _tier_name,
                "card": {k: v for k, v in c.items() if not str(k).startswith("_")},
                "first_tracked": prev.get("first_tracked", today_iso),
                "last_confirmed": today_iso
            }

    live_ids = set(tracked.keys()) & ({c.get("entry_id") for c in urgent if c.get("entry_id")} |
                                       {c.get("entry_id") for c in needs if c.get("entry_id")})

    # 2. Resolve or carry anything that scrolled out of this run's pull.
    carried = 0
    dropped_resolved = 0
    inconclusive = 0
    stale_warnings = 0
    for eid, rec in list(tracked.items()):
        if eid in live_ids:
            continue  # already handled by the checkpoint above
        if eid in _cc_done_entry_ids:
            del tracked[eid]
            dropped_resolved += 1
            continue
        if eid in _ticked_done_entry_ids:
            del tracked[eid]
            dropped_resolved += 1
            continue

        outcome = "unknown"
        try:
            item = mapi.GetItemFromID(eid)
            item_parent_id = item.Parent.EntryID
            outcome = "still_open" if item_parent_id == _inbox_folder.EntryID else "moved_out"
        except Exception:
            outcome = "unknown"

        if outcome == "moved_out":
            del tracked[eid]
            dropped_resolved += 1
            continue
        if outcome == "unknown":
            inconclusive += 1
            # fall through to carry -- fail open, see comment block above

        card_copy = dict(rec.get("card") or {})
        if not card_copy:
            del tracked[eid]
            continue
        card_copy["_carried_forward"] = True
        tier = rec.get("tier", "needs")
        (urgent if tier == "urgent" else needs).append(card_copy)
        tracked[eid]["last_confirmed"] = today_iso
        carried += 1

        try:
            _first = datetime.strptime(rec.get("first_tracked", today_iso), "%Y-%m-%d")
            age_days = (datetime.now() - _first).days
            if age_days > 90:
                stale_warnings += 1
                print(f"WARNING: Phase 3.9 - '{card_copy.get('subject','?')}' has been carried for {age_days} days without resolving (still shown, not dropped)")
        except Exception:
            pass

    if carried or dropped_resolved or inconclusive:
        dry_run_tag = " [DRY RUN - no writes]" if _WI_PHASE39_DRY_RUN else ""
        print(f"Phase 3.9 done - carried:{carried} dropped_resolved:{dropped_resolved} inconclusive_lookups_carried:{inconclusive} stale_over_90d:{stale_warnings} tracked_total:{len(tracked)}{dry_run_tag}")

    if _WI_PHASE39_DRY_RUN:
        print("Phase 3.9 dry run - skipping triage_ledger.json write and briefing carry-forward injection."
              + ("  [WI_AI_PARALLEL]" if AI_PARALLEL else ""))
        # Dry run intentionally still ran the resolution-signal checks above
        # for real (Outlook/CC/ticks), it just doesn't persist the ledger or
        # leave carried cards appended to urgent/needs -- undo the in-memory
        # append so a dry run has zero observable effect on this run's output.
        if carried:
            urgent[:] = [c for c in urgent if not c.get("_carried_forward")]
            needs[:] = [c for c in needs if not c.get("_carried_forward")]
    elif GITHUB_PAT:
        try:
            _persist_ledger["tracked_needs_urgent"] = tracked
            _persist_rw_headers = {"Authorization": f"token {GITHUB_PAT}", "Content-Type": "application/json", "User-Agent": "work-inbox-script"}
            _gh_put(_persist_ledger_url, _persist_rw_headers,
                    "chore: update Needs/Urgent scroll-out tracking",
                    json.dumps(_persist_ledger, indent=1).encode("utf-8"), _persist_sha)
        except Exception as e:
            print(f"WARNING: Phase 3.9 could not persist triage_ledger.json - {e}")
except Exception as e:
    print(f"WARNING: Phase 3.9 scroll-out persistence failed entirely, Urgent/Needs left as fresh-pull-only this run - {e}")

briefing = {
    "date":         today_str,
    "subtitle":     subtitle,
    "context":      context,
    "urgent":       urgent,
    "needs":        needs,
    "fyi":          fyi,
    "fyiRawCount":  fyi_raw_count,
    "low":          low,
    "calToday":     cal_today_items,
    "calTomorrow":  cal_tomorrow_items,
    "calDay2":      cal_day2_items,
    "calDay3":      cal_day3_items,
    "calFull":      calFull,
    "absences":     absences,
    "prioritiesToday":    priorities_today,
    "prioritiesTomorrow": priorities_tomorrow,
    "prioritiesWeek":     priorities_week,
    "refreshed_at": datetime.now().strftime("%A %d %B · %H:%M")
}

# Laptop bridge: make the missing calendar explicit rather than silently absent.
# (Absences can still be present here -- Phase "Absence preservation" above carries
# the last full briefing's same-day absences forward -- so the note is about the
# meeting list only.)
if BRIDGE_ALLOW_EMPTY_CALENDAR and calendar_summary_count(briefing) == 0:
    briefing["calendarUnavailable"] = True
    _carried = (" Colleague absences below are carried forward from the last full briefing."
                if briefing["absences"] else "")
    _bridge_note = ("Calendar unavailable this run (bridge mode - the laptop has no calendar "
                    "source): no meeting list or prep for today or this week." + _carried)
    briefing["context"] = ((context + " " + _bridge_note).strip() if context else _bridge_note)
    log("Phase 4 - bridge mode: no calendar summaries; set calendarUnavailable + noted it in context.")

# -- Phase 4 -- push to GitHub --
log("Phase 4 - pushing briefing to GitHub...")

if AI_PARALLEL:
    _parallel_briefing_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "data", "claude_briefing.json")
    with open(_parallel_briefing_path, "w", encoding="utf-8") as _pf:
        json.dump(briefing, _pf, indent=2, ensure_ascii=False)
    log(f"Phase 4 - PARALLEL VALIDATION MODE: wrote {_parallel_briefing_path} locally, pushed NOTHING to GitHub.")
elif not GITHUB_PAT:
    print("ERROR: GITHUB_PAT env var not set - cannot push.")
else:
    try:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_PATH}"
        headers = {
            "Authorization": f"token {GITHUB_PAT}",
            "Content-Type":  "application/json",
            "User-Agent":    "work-inbox-script"
        }
        sha = None
        remote_meta = None
        remote_briefing = existing_briefing
        try:
            remote_meta = _gh_get(api_url, headers)
            sha = remote_meta.get("sha")
            if remote_meta.get("content"):
                remote_bytes = base64.b64decode(remote_meta["content"])
                remote_briefing = json.loads(remote_bytes.decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise

        fatal, warnings = validate_briefing_update(
            briefing, remote_briefing,
            allow_empty_calendar=BRIDGE_ALLOW_EMPTY_CALENDAR)
        for warning in warnings:
            print(f"Phase 4 safe-write warning - {warning}")
        if fatal:
            raise Exception("Safe write blocked briefing update: " + "; ".join(fatal))

        _backup_briefing_before_write(remote_meta, headers)

        content_b64 = base64.b64encode(
            json.dumps(briefing, indent=2, ensure_ascii=False).encode("utf-8")
        ).decode("ascii")
        payload = {
            "message": f"chore: update briefing {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "content": content_b64
        }
        if sha:
            payload["sha"] = sha
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(api_url, data=data, headers=headers, method="PUT")
        with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
            result = json.loads(r.read())
            print(f"Phase 4 done - briefing pushed to GitHub (commit: {result.get('commit',{}).get('sha','?')[:7]})")
    except Exception as e:
        print(f"Phase 4 FAILED - {e}")
        raise


# Phase 5 - push task suggestions to GitHub (consumed by Command Centre dashboard)
if AI_PARALLEL:
    _parallel_sug_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "data", "claude_inbox_suggestions.json")
    with open(_parallel_sug_path, "w", encoding="utf-8") as _psf:
        json.dump(suggestions, _psf, indent=2, ensure_ascii=False)
    log(f"Phase 5 - PARALLEL VALIDATION MODE: wrote {_parallel_sug_path} locally, pushed NOTHING (no carry-forward from remote).")
elif GITHUB_PAT:
    try:
        sug_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/inbox_suggestions.json"
        headers = {
            "Authorization": f"token {GITHUB_PAT}",
            "Content-Type":  "application/json",
            "User-Agent":    "work-inbox-script"
        }
        sha = None
        prev_suggestions = None
        try:
            req = urllib.request.Request(sug_url, headers=headers)
            with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
                _prev_meta = json.loads(r.read())
                sha = _prev_meta.get("sha")
                prev_suggestions = json.loads(base64.b64decode(_prev_meta["content"]).decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
        except Exception:
            prev_suggestions = None

        # Carry forward any earlier suggestion that never became a task, so an
        # unactioned suggestion is not silently lost when this file is
        # rewritten (the script now runs five times a day).
        if prev_suggestions:
            seen = {s.get("entry_id") for s in suggestions["new_tasks"] if s.get("entry_id")}
            carried = 0
            carry_suppressed_no_action = 0
            for old in prev_suggestions.get("new_tasks", []):
                oid = old.get("entry_id")
                if not oid or oid in seen or oid in ledger.get("promoted", {}):
                    continue
                # 12 Aug 2026: also drop a previously-persisted suggestion
                # here if its source email was AI-confirmed no_action_needed
                # by THIS run's Phase 3.3/3.3b (demoted out of Needs/Urgent
                # just above) -- Codex review flagged that without this
                # check, a noisy suggestion generated earlier could keep
                # resurfacing via carry-forward even after the fresh Phase
                # 3.5 output above was correctly filtered in the same run.
                # Known limitation, not fixed here: `_noise_demoted_entry_ids`
                # is process-local to this run only (line ~721) -- it has no
                # memory of a PAST run's demotions, so an old carried-forward
                # suggestion whose source email has since scrolled out of the
                # 50-newest-email inbox window (so it's no longer in this
                # run's summary_candidates at all) won't be caught here. A
                # full fix would need to persist demoted entry_ids across
                # runs (e.g. in triage_ledger.json) -- out of scope for this
                # extension, flagged not silently dropped.
                if oid in _noise_demoted_entry_ids:
                    carry_suppressed_no_action += 1
                    continue
                old["carried_forward"] = True
                suggestions["new_tasks"].append(old)
                seen.add(oid)
                carried += 1
            if carried or carry_suppressed_no_action:
                print(f"Phase 5 - carried forward {carried} unactioned suggestion(s) (suppressed_no_action:{carry_suppressed_no_action})")

        content_b64 = base64.b64encode(
            json.dumps(suggestions, indent=2, ensure_ascii=False).encode("utf-8")
        ).decode("ascii")
        payload = {
            "message": f"chore: update task suggestions {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "content": content_b64
        }
        if sha:
            payload["sha"] = sha
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(sug_url, data=data, headers=headers, method="PUT")
        with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
            result = json.loads(r.read())
            print(f"Phase 5 done - suggestions pushed (commit: {result.get('commit',{}).get('sha','?')[:7]})")
    except Exception as e:
        print(f"Phase 5 FAILED - {e}")
