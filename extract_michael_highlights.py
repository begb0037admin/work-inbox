"""
extract_michael_highlights.py
Reads michael_evidence.json, extracts key sentences per category,
and pushes 4 separate text files to GitHub (one per category group).

Run:
    python extract_michael_highlights.py
"""

import json, os, base64, re, urllib.request, urllib.error, time
from datetime import datetime

GITHUB_REPO = "begb0037admin/work-inbox"
GITHUB_PAT  = os.environ.get("GITHUB_PAT", "")

# 4 groups — one output file each
GROUPS = {
    "evidence_1_tupe_ndm_osps_balance": {
        "TUPE / Liverpool School of Medicine": [
            "tupe", "liverpool", "transfer of undertaking",
        ],
        "NDM Workgroup Review": [
            "ndm", "workgroup", "work group",
        ],
        "OSPS Pension Contribution Rate": [
            "osps", "contribution rate", "pension contribution", "1oct26",
        ],
        "Balance Maintenance / TUPE Balances": [
            "balance maintenance", "balance transfer", "balance migration",
            "tupe balance",
        ],
    },
    "evidence_2_absence_college_ess_ref": {
        "Sickness Absence Code / Allow Employee Requests": [
            "allow employee request", "absence request", "absence code",
            "non sickness", "non-sickness", "absence restructure",
            "sickness code", "timdep",
        ],
        "College Staff / DTP1092 / REF2029": [
            "dtp1092", "dtp 1092", "college staff", "college data",
            "pay admin", "ref2029", "ref 2029", "pay administered",
        ],
        "Employee Self Service / Portal": [
            "employee self service", "ess ", "self service portal",
            "portal migration", "back office to portal",
        ],
    },
    "evidence_3_roster_timesheet_holiday": {
        "Rostering / Roster": [
            "roster", "rostering", "glam roster", "sbs roster",
            "unrostered", "variable hr",
        ],
        "Timesheet Pay Award Arrears": [
            "pay award arrear", "timesheet arrear", "arrear",
            "pay award", "timesheet pay",
        ],
        "WFM Holiday / Rolled Up Pay": [
            "rolled up hol", "rolled-up hol", "holiday pay accrual",
            "wfm holiday period", "holiday period end",
        ],
        "Volunteering Hours": [
            "volunteer", "volunteering hours",
        ],
    },
    "evidence_4_praise_thanks": {
        "Praise / Thanks for Michael": [
            "thank you michael", "thanks michael", "well done michael",
            "michael has", "michael is", "michael was", "michael did",
            "michael always", "michael never", "credit to michael",
            "michael's work", "michael's contribution",
            "michael o'sullivan",
        ],
    },
}

EXCLUDE_SUBJ = [
    "clinical pay", "ucea", "clinical uplift",
    "github", "pages build", "workflow", "deploy",
    "out of office", "automatic reply", "accepted:", "declined:",
    "newsletter", "digest",
]

def sentences_around(text, keyword, window=2):
    sents = re.split(r'(?<=[.!?])\s+', text)
    hits = []
    for i, s in enumerate(sents):
        if keyword.lower() in s.lower():
            start = max(0, i - 1)
            end   = min(len(sents), i + window)
            excerpt = " ".join(sents[start:end]).strip()
            if len(excerpt) > 30:
                hits.append(excerpt[:400])
    return hits[:2]

# Fetch full evidence file once
print("Fetching michael_evidence.json...")
t = int(time.time())
raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/data/michael_evidence.json?t={t}"
req = urllib.request.Request(raw_url, headers={"User-Agent": "extract-script"})
with urllib.request.urlopen(req) as r:
    data = json.loads(r.read().decode("utf-8"))
emails = data.get("emails", [])
print(f"Loaded {len(emails)} emails")

def push_file(path, text):
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
    headers = {
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
    payload = {
        "message": f"chore: michael evidence {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "content": base64.b64encode(text.encode("utf-8")).decode("ascii"),
    }
    if sha:
        payload["sha"] = sha
    req2 = urllib.request.Request(api_url,
                                   data=json.dumps(payload).encode("utf-8"),
                                   headers=headers, method="PUT")
    with urllib.request.urlopen(req2) as r:
        result = json.loads(r.read())
        return result.get("commit", {}).get("sha", "?")[:7]

for group_name, categories in GROUPS.items():
    results = {cat: [] for cat in categories}

    for email in emails:
        subj = (email.get("subject") or "")
        body = (email.get("body") or "")
        date = (email.get("date") or "")[:10]
        frm  = email.get("from") or ""

        if any(ex in subj.lower() for ex in EXCLUDE_SUBJ):
            continue

        for cat, terms in categories.items():
            for term in terms:
                if term in subj.lower() or term in body.lower():
                    excerpts = sentences_around(body, term)
                    if not excerpts:
                        excerpts = [body[:300].strip()]
                    hit = {"subject": subj, "from": frm, "date": date, "excerpts": excerpts}
                    if not any(h["subject"] == subj and h["date"] == date for h in results[cat]):
                        results[cat].append(hit)
                    break

    lines = [
        f"MICHAEL O'SULLIVAN — EXCELLENCE AWARD EVIDENCE",
        f"Group: {group_name}",
        f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}",
        "=" * 60,
    ]
    for cat, hits in results.items():
        if not hits:
            lines.append(f"\n[{cat}]\n  -- No emails found --")
            continue
        lines.append(f"\n[{cat}] ({len(hits)} emails)")
        lines.append("-" * 50)
        for h in hits[:8]:
            lines.append(f"  Subject: {h['subject']}")
            lines.append(f"  From: {h['from']}  |  Date: {h['date']}")
            for exc in h["excerpts"]:
                lines.append(f"  >> {exc}")
            lines.append("")

    text = "\n".join(lines)
    out_path = f"data/{group_name}.txt"
    commit = push_file(out_path, text)
    print(f"Pushed {out_path} ({len(text)//1000}KB) - commit {commit}")

print("\nAll 4 files pushed. Claude can now read them.")
