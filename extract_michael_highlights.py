"""
extract_michael_highlights.py
Reads the full michael_evidence.json, extracts the most relevant sentences
around key evidence terms, and pushes a compact readable text file to GitHub.

Run:
    python extract_michael_highlights.py
"""

import json, os, base64, re, urllib.request, urllib.error, time
from datetime import datetime

GITHUB_REPO = "begb0037admin/work-inbox"
GITHUB_PAT  = os.environ.get("GITHUB_PAT", "")

# Fetch full evidence file
print("Fetching michael_evidence.json...")
t = int(time.time())
raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/data/michael_evidence.json?t={t}"
req = urllib.request.Request(raw_url, headers={"User-Agent": "extract-script"})
with urllib.request.urlopen(req) as r:
    data = json.loads(r.read().decode("utf-8"))

emails = data.get("emails", [])
print(f"Loaded {len(emails)} emails")

# Evidence categories to find
CATEGORIES = {
    "TUPE / Liverpool School of Medicine": [
        "tupe", "liverpool", "transfer of undertaking",
    ],
    "NDM Workgroup Review": [
        "ndm", "workgroup", "work group", "ndm workgroup",
    ],
    "Sickness Absence Code / Allow Employee Requests": [
        "allow employee request", "absence request", "absence code",
        "non sickness", "non-sickness", "absence restructure",
        "sickness code", "timdep",
    ],
    "OSPS Pension Contribution Rate": [
        "osps", "contribution rate", "pension contribution", "1oct26",
    ],
    "College Staff / DTP1092": [
        "dtp1092", "dtp 1092", "college staff", "college data",
        "pay admin", "ref2029", "ref 2029", "pay administered",
    ],
    "Rostering / Roster": [
        "roster", "rostering", "glam roster", "sbs roster",
        "unrostered", "variable hr", "12 month",
    ],
    "Timesheet Pay Award Arrears": [
        "pay award arrear", "timesheet arrear", "arrear",
        "pay award", "timesheet pay",
    ],
    "Balance Maintenance / TUPE Balances": [
        "balance maintenance", "balance transfer", "balance migration",
        "tupe balance",
    ],
    "Employee Self Service / Portal": [
        "employee self service", "ess ", "self service portal",
        "portal migration", "back office to portal",
    ],
    "Praise / Thanks for Michael": [
        "thank you michael", "thanks michael", "well done michael",
        "michael has", "michael is", "michael was", "michael did",
        "michael always", "michael never", "credit to michael",
        "michael's work", "michael's contribution",
    ],
    "Volunteering Hours": [
        "volunteer", "volunteering hours",
    ],
    "WFM Holiday / Rolled Up Pay": [
        "rolled up hol", "rolled-up hol", "holiday pay accrual",
        "wfm holiday period", "holiday period end",
    ],
}

EXCLUDE_SUBJ = [
    "clinical pay", "ucea", "clinical uplift", "clinical pay uplift",
    "github", "pages build", "workflow", "deploy", "out of office",
    "automatic reply", "accepted:", "declined:", "newsletter",
]

def sentences_around(text, keyword, window=2):
    """Return up to `window` sentences either side of the keyword match."""
    sents = re.split(r'(?<=[.!?])\s+', text)
    hits = []
    for i, s in enumerate(sents):
        if keyword.lower() in s.lower():
            start = max(0, i - 1)
            end   = min(len(sents), i + window)
            excerpt = " ".join(sents[start:end]).strip()
            if len(excerpt) > 30:
                hits.append(excerpt[:400])
    return hits[:2]  # max 2 excerpts per keyword per email

results = {}  # category -> list of hits

for cat, terms in CATEGORIES.items():
    results[cat] = []

for email in emails:
    subj = (email.get("subject") or "")
    body = (email.get("body") or "")
    date = (email.get("date") or "")[:10]
    frm  = email.get("from") or ""

    subj_lower = subj.lower()
    if any(ex in subj_lower for ex in EXCLUDE_SUBJ):
        continue

    for cat, terms in CATEGORIES.items():
        for term in terms:
            if term in subj_lower or term in body.lower():
                excerpts = sentences_around(body, term)
                if not excerpts:
                    excerpts = [body[:300].strip()]
                hit = {
                    "subject":  subj,
                    "from":     frm,
                    "date":     date,
                    "excerpts": excerpts,
                }
                # Dedupe by subject+date
                existing = [h for h in results[cat]
                            if h["subject"] == subj and h["date"] == date]
                if not existing:
                    results[cat].append(hit)
                break  # one hit per email per category

# Build readable text output
lines = []
lines.append("MICHAEL O'SULLIVAN — EXCELLENCE AWARD EVIDENCE")
lines.append(f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}")
lines.append("=" * 60)

for cat, hits in results.items():
    if not hits:
        lines.append(f"\n[{cat}]\n  -- No emails found --")
        continue
    lines.append(f"\n[{cat}] ({len(hits)} emails)")
    lines.append("-" * 50)
    for h in hits[:6]:  # max 6 per category
        lines.append(f"  Subject: {h['subject']}")
        lines.append(f"  From: {h['from']}  |  Date: {h['date']}")
        for exc in h["excerpts"]:
            lines.append(f"  >> {exc}")
        lines.append("")

output_text = "\n".join(lines)

# Push to GitHub as plain text
out_path = "data/michael_award_evidence.txt"
api_url  = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{out_path}"
headers  = {
    "Authorization": f"token {GITHUB_PAT}",
    "Content-Type":  "application/json",
    "User-Agent":    "extract-michael-highlights",
}

sha = None
try:
    r2 = urllib.request.Request(api_url, headers=headers)
    with urllib.request.urlopen(r2) as r:
        sha = json.loads(r.read()).get("sha")
except urllib.error.HTTPError as e:
    if e.code != 404:
        raise

content_b64 = base64.b64encode(output_text.encode("utf-8")).decode("ascii")
payload = {"message": f"chore: michael award evidence {datetime.now().strftime('%Y-%m-%d %H:%M')}",
           "content": content_b64}
if sha:
    payload["sha"] = sha

req2 = urllib.request.Request(api_url,
                               data=json.dumps(payload).encode("utf-8"),
                               headers=headers, method="PUT")
with urllib.request.urlopen(req2) as r:
    result = json.loads(r.read())
    commit = result.get("commit", {}).get("sha", "?")[:7]
    print(f"Done - commit {commit}")
    print(f"Pushed: {GITHUB_REPO}/{out_path}")
    char_count = len(output_text)
    print(f"Output size: {char_count} chars (~{char_count//1000}KB)")

# Also print a preview locally
print("\n--- PREVIEW ---")
print(output_text[:2000])
