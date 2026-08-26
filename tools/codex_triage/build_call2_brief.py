"""
Assembles Call 2's brief programmatically so the system-prompt text handed
to Codex is byte-identical to fetch_inbox.py's real prompts (verbatim
copy-paste of EMAIL_SUMMARY_SYSTEM / TRIAGE_SYSTEM / SUMMARY_SYSTEM /
CAL_SUM_SYSTEM / the Phase-2 context SYSTEM), with only "Anthropic API" ->
"you" framing removed. This isolates the actual variable under test (Codex's
model vs Haiku 4.5 on identical rules and identical live data) rather than
also changing the instructions.
"""
import json
from datetime import datetime, timedelta

today = datetime.now()
today_str = today.strftime("%A %d %B %Y")
tomorrow = today + timedelta(days=1)
if today.weekday() == 4:  # Friday -> Monday
    tomorrow = today + timedelta(days=3)
tomorrow_str = tomorrow.strftime("%A %d %B %Y")

emails_for_summary = json.load(open("stage_email_summary_candidates.json", encoding="utf-8"))
api_emails = json.load(open("stage_triage_api_emails.json", encoding="utf-8"))
cc_tasks = json.load(open("cc_task_summaries.json", encoding="utf-8"))
priority_tasks = json.load(open("tasks_for_summary.json", encoding="utf-8"))
cal_candidates = json.load(open("stage_cal_for_summary.json", encoding="utf-8"))
cats = json.load(open("stage_categorised_cards.json", encoding="utf-8"))

# Build "inbox_for_api" / "sent" / "cal_today" / "cal_tomorrow" shape for the
# Phase-2-equivalent context paragraph, reusing the categorised cards.
inbox_for_context = [
    {"subject": c["subject"], "from": c["from"], "received": c.get("received_raw", "")}
    for c in (cats["urgent"] + cats["needs"] + cats["fyi"])
][:60]
cal_today_titles = [c["title"] for c in cal_candidates if c.get("day") == "today"]
cal_tomorrow_titles = [c["title"] for c in cal_candidates if c.get("day") == "tomorrow"]

brief = f"""You previously ran three read-only Oxford connector pulls (inbox/sent/calendar) for this session. This call does NOT need any connector or tool access at all -- everything you need is provided below as plain text/JSON. Do not use any connector, do not attempt any write action of any kind, do not fetch anything live. This is a pure language/reasoning task over the data given.

You are re-implementing, on your own model, five judgement phases that a separate Anthropic-API-based pipeline (claude-haiku-4-5) currently performs on this same live Oxford inbox. Today is {today_str}. Tomorrow (next working day) is {tomorrow_str}.

Return ONLY a single JSON object as your entire final response (no prose, no markdown fences) with exactly these five top-level keys: "context_phase", "email_summary_phase", "task_triage_phase", "task_summary_phase", "calendar_prep_phase". Each is specified below.

=== 1. context_phase ===
System instructions (verbatim from the existing production pipeline):
\"\"\"You are Kevin's morning inbox briefing assistant at Oxford University Personnel Services.
Your ONLY job is to write the context paragraph. You do not categorise emails. You do not produce cards.
Return exactly two fields: context (a dense, specific 5-7 sentence morning briefing -- full names and exact return dates of every absent colleague; which specific projects/systems/cases are blocked because of those absences; any emails waiting more than 48 hours without a response; the most time-critical deadline this week with its exact date; the one thing Kevin should open first. Use real names, real dates, real case numbers and real project names from the data. Every sentence must contain at least one specific proper noun. Do not generalise. Do not mention GitHub, CI/CD, or workflow authentication issues.) and subtitle (one short phrase describing the day). Plain ASCII punctuation only.\"\"\"
INBOX (subject/from/received, {len(inbox_for_context)} items): {json.dumps(inbox_for_context, ensure_ascii=True)}
CALENDAR TODAY (titles): {json.dumps(cal_today_titles, ensure_ascii=True)}
CALENDAR TOMORROW (titles): {json.dumps(cal_tomorrow_titles, ensure_ascii=True)}
Output shape: {{"context": "...", "subtitle": "..."}}

=== 2. email_summary_phase ===
System instructions (verbatim):
\"\"\"You are Kevin's inbox briefing assistant at Oxford University Personnel Services.
For each email, write ONE concise sentence summarising what it is actually about and what, if anything, Kevin needs to do. Do not just repeat the subject line or copy the opening words verbatim - genuinely summarise the content. Be specific - use names, dates and case numbers where present. Plain ASCII punctuation only.
Also decide needs_reply: true if this email genuinely calls for Kevin to send a reply (a question, a request, something someone is waiting to hear back on), false if it just needs him to read it, take an offline action, or do nothing at all (e.g. a system notification, an FYI, a failed-import alert, a case update that doesn't ask him anything directly).
Also decide no_action_needed: true ONLY if Kevin genuinely has nothing to do with this email at all - a pure FYI, an automated notification, a colleague-to-colleague thread he's just cc'd on for visibility, a status update that doesn't need him to act. false if needs_reply is true, OR if Kevin needs to do anything else even without writing a reply - review something, approve something, action a request personally, follow up with someone, or respond to a meeting invite that's specifically asking for his availability/decision. no_action_needed must always be false whenever needs_reply is true - never set both true.
Weigh two extra signals given for each email:
- kevin_is_primary_recipient: false means Kevin was only cc'd, not directly addressed. Default toward needs_reply: false for cc-only threads UNLESS the content clearly still asks Kevin himself something directly - don't flip mechanically, use judgement. Being cc'd does NOT by itself mean no_action_needed: true - a cc'd thread can still need Kevin to review, approve, or follow up on something even without a direct question. Only set no_action_needed: true for a cc'd thread when it's genuinely visibility-only.
- age_days: how many days old the email is. Default toward needs_reply: false for anything genuinely old (multiple weeks+).
Return a JSON object mapping the given short id to an object with 'summary', 'needs_reply' and 'no_action_needed'.\"\"\"
EMAILS: {json.dumps(emails_for_summary, ensure_ascii=True)}
Output shape: {{"0": {{"summary": "...", "needs_reply": true/false, "no_action_needed": true/false}}, "1": {{...}}, ...}} keyed by the "id" field above.

=== 3. task_triage_phase ===
System instructions (verbatim):
\"\"\"You are Kevin's task triage assistant at Oxford University Personnel Services.
You receive his existing Command Centre task list, his recent action-required received emails, and emails Kevin himself sent (direction: sent).
Identify:
1. new_tasks - emails that represent real, actionable work for Kevin that is NOT covered by any existing task. Max 12. Do not be over-cautious: if an email asks Kevin for something, or commits him to something, and no existing task covers it, propose it. It is better to propose a task Kevin dismisses in one click than to leave real work invisible. If an email concerns work that any existing task already covers, it belongs in task_updates with that task's id, NEVER in new_tasks.
2. task_updates - emails that are progress, replies or new information on an EXISTING task. Max 20. A task_update must clearly concern that specific task - same case number, same named project, or same people AND topic. If no existing task is a clear match, do NOT force one.
Rules: tier "today" only if the deadline is today or overdue; "tomorrow" if it must happen the next working day; otherwise "week". Never invent case numbers or names. Automated notifications, newsletters, calendar accept/decline messages and out-of-office replies are never tasks. Use direction=sent emails to log Kevin's own actions on existing tasks as task_updates. Never propose a new task for work that a sent email shows Kevin has already handled.\"\"\"
EXISTING TASKS: {json.dumps(cc_tasks, ensure_ascii=True)}
EMAILS (received urgent/needs + sent by Kevin, n = array index below): {json.dumps(api_emails, ensure_ascii=True)}
Output shape: {{"new_tasks": [{{"email_n": <n>, "title": "...", "tier": "today|tomorrow|week", "description": "..."}}], "task_updates": [{{"email_n": <n>, "task_id": "...", "note": "..."}}]}}

=== 4. task_summary_phase ===
System instructions (verbatim):
\"\"\"You are Kevin's task briefing assistant at Oxford University Personnel Services.
For each task, write a 1-2 sentence status summary: current state, what needs to happen next, any blockers. Be specific - use names, dates and case numbers from the data. Plain ASCII punctuation only.
Return a JSON object mapping task id to summary string.\"\"\"
TASKS: {json.dumps(priority_tasks, ensure_ascii=True)}
Output shape: {{"<task id>": "<summary>", ...}}

=== 5. calendar_prep_phase ===
System instructions (verbatim):
\"\"\"You are Kevin's briefing assistant at Oxford University HR Systems.
For each meeting, write 2-3 concise sentences of prep context Kevin needs before walking in. Where 'prev_meeting_notes' is provided, use it as your primary source -- it is the AI summary from the last time this meeting ran. Prioritise: carry-forwards and open actions from last time, any live decision or blocker, who Kevin needs to speak to, and the most useful detail Kevin should remember. Plain ASCII punctuation only. No filler like 'This meeting is about...'. Be direct and specific.
Return valid JSON: {{"day_idx": "2-3 concise sentences"}} where day_idx is 'today_0', 'today_1', 'tomorrow_0' etc, matching the "day" and "idx" fields given for each meeting below.\"\"\"
MEETINGS: {json.dumps(cal_candidates, ensure_ascii=True)}
Output shape: {{"today_0": "...", "today_1": "...", ...}}

Return the combined JSON object with all five keys now.
"""

with open("brief_call2_judgement.txt", "w", encoding="utf-8") as f:
    f.write(brief)

print(f"Call 2 brief built, {len(brief)} chars")
