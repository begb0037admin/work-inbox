"""
filter_michael_evidence.py
Reads michael_evidence.json from GitHub, strips out WFM/clinical pay emails
(already covered), and pushes a compact highlights file Claude can read.

Run:
    python filter_michael_evidence.py
"""

import json, os, base64, urllib.request, urllib.error, time
from datetime import datetime

GITHUB_REPO = "begb0037admin/work-inbox"
GITHUB_PAT  = os.environ.get("GITHUB_PAT", "")

# --- Fetch full evidence file from GitHub ---
print("Fetching michael_evidence.json from GitHub...")
t = int(time.time())
raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/data/michael_evidence.json?t={t}"
req = urllib.request.Request(raw_url, headers={"User-Agent": "filter-script"})
with urllib.request.urlopen(req) as r:
    data = json.loads(r.read().decode("utf-8"))

emails = data.get("emails", [])
print(f"Loaded {len(emails)} emails")

# --- Exclusion: subjects already well-evidenced ---
EXCLUDE_SUBJECTS = [
    "clinical pay", "ucea", "clinical uplift",
    "wfm wave 3", "wave 3", "wfm l&a wave",
    "leave & absence wave", "leave and absence wave",
    "github", "pages build", "workflow", "deploy",
    "out of office", "automatic reply", "accepted:", "declined:",
    "newsletter", "digest", "unsubscribe",
]

# --- Target: subjects/bodies we want for the nomination ---
TARGET_TERMS = [
    "tupe", "liverpool",
    "ndm", "workgroup", "work group",
    "osps", "contribution rate", "pension contribution",
    "sickness absence", "absence code", "absence request", "non sickness", "non-sickness",
    "allow employee", "employee request",
    "college staff", "dtp1092",
    "calendar view", "employee calendar",
    "self service", "ess ",
    "ref2029", "ref 2029", "pay admin",
    "rostering", "roster",
    "volunteering", "volunteer hours",
    "rolled up", "rolled-up", "holiday pay accrual",
    "variable hours", "unrostered",
    "michael o'sullivan", "michael o sullivan",
    "well done", "thank you michael", "thanks michael",
    "great work", "excellent", "outstanding", "praise",
    "timesheets", "timesheet pay", "pay award arrears",
    "balance maintenance", "balance transfer",
]

def is_excluded(email):
    subj = (email.get("subject") or "").lower()
    return any(ex in subj for ex in EXCLUDE_SUBJECTS)

def is_target(email):
    subj = (email.get("subject") or "").lower()
    body = (email.get("body") or "").lower()[:2000]
    return any(t in subj or t in body for t in TARGET_TERMS)

filtered = [e for e in emails if not is_excluded(e) and is_target(e)]
print(f"After filtering: {len(filtered)} relevant emails")

# Trim bodies to 1500 chars to keep output small
for e in filtered:
    if e.get("body"):
        e["body"] = e["body"][:1500]

output = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "total_filtered": len(filtered),
    "emails": filtered,
}

# --- Push compact file to GitHub ---
out_path = "data/michael_evidence_highlights.json"
api_url  = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{out_path}"
headers  = {
    "Authorization": f"token {GITHUB_PAT}",
    "Content-Type":  "application/json",
    "User-Agent":    "filter-michael-evidence",
}

sha = None
try:
    req = urllib.request.Request(api_url, headers=headers)
    with urllib.request.urlopen(req) as r:
        sha = json.loads(r.read()).get("sha")
except urllib.error.HTTPError as e:
    if e.code != 404:
        raise

content_b64 = base64.b64encode(
    json.dumps(output, indent=2, ensure_ascii=False).encode("utf-8")
).decode("ascii")
payload = {
    "message": f"chore: michael evidence highlights {datetime.now().strftime('%Y-%m-%d %H:%M')}",
    "content": content_b64,
}
if sha:
    payload["sha"] = sha

req = urllib.request.Request(api_url,
                              data=json.dumps(payload).encode("utf-8"),
                              headers=headers, method="PUT")
with urllib.request.urlopen(req) as r:
    result = json.loads(r.read())
    commit = result.get("commit", {}).get("sha", "?")[:7]
    print(f"Done - commit {commit}")
    print(f"Pushed: {GITHUB_REPO}/{out_path}")
    print(f"Emails in highlights: {len(filtered)}")
