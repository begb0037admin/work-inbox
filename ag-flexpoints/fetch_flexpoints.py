import json, os, base64, re, glob, urllib.request, urllib.error
from datetime import datetime
import anthropic

# AG FlexPoints - recurring report generator
# Runs locally (admin Windows machine), same pattern as fetch_inbox.py:
#   Phase 1: pull FlexPoints pages from Access Customer Success Portal
#            (login via ACCESS_PORTAL_USER / ACCESS_PORTAL_PASSWORD env vars;
#             falls back to any page saved manually into source/)
#   Phase 2: Anthropic API (claude-haiku-4-5) parses pages -> structured report
#   Phase 3: snapshot appended to history, latest report written locally
#   Phase 4: push report + history to GitHub via Contents API

GITHUB_REPO  = "begb0037admin/AG-FlexPoints"
REPORT_PATH  = "data/flexpoints.json"
HISTORY_PATH = "data/flexpoints_history.json"
GITHUB_PAT   = os.environ.get("GITHUB_PAT", "")

PORTAL_BASE  = "https://accessgroup.my.site.com/Support"
LOGIN_URL    = PORTAL_BASE + "/s/login/"
PORTAL_PAGES = {
    "flexpoints":   PORTAL_BASE + "/s/flexpoints",
    "our_services": PORTAL_BASE + "/s/our-services",
}

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_DIR = os.path.join(SCRIPT_DIR, "source")
MAX_DETAIL_PAGES = 25  # detail pages visited per run (newest items first)

# -- Phase 1 - pull portal pages ---------------------------------------------

REF_PATTERN = re.compile(r"(FPT-\d+|\d{6,10})$")

def harvest_links(anchors, links):
    """Collect ref -> portal URL so every dashboard/PDF item links to its
    exact transaction detail page."""
    for a in anchors:
        text = (a.get("text") or "").strip()
        href = (a.get("href") or "").strip()
        if href and REF_PATTERN.fullmatch(text):
            if href.startswith("/"):
                href = "https://accessgroup.my.site.com" + href
            links[text] = href

def get_portal_text():
    """Log in to the Access portal with Playwright and return rendered page text.
    Credentials come from Windows User env vars - never stored in any file."""
    user = os.environ.get("ACCESS_PORTAL_USER", "")
    pwd  = os.environ.get("ACCESS_PORTAL_PASSWORD", "")
    if not user or not pwd:
        print("Phase 1 - ACCESS_PORTAL_USER / ACCESS_PORTAL_PASSWORD not set, skipping live pull")
        return None
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Phase 1 - playwright not installed (pip install playwright && playwright install chromium)")
        return None
    USER_SELECTORS = ('input[type="email"]', 'input[name*="user" i]',
                      'input[id*="user" i]', 'input[autocomplete="username"]',
                      'input[placeholder*="sername" i]', 'input[placeholder*="mail" i]',
                      'lightning-input input', 'input[type="text"]')
    PASS_SELECTORS = ('input[type="password"]', 'input[autocomplete="current-password"]')
    SUBMIT_SELECTORS = ('button:has-text("Log in")', 'button:has-text("Login")',
                        'button:has-text("Sign in")', 'button[type="submit"]',
                        'input[type="submit"]')

    def find_in_frames(page, selectors):
        """Salesforce login forms often live inside an iframe - search all frames."""
        for frame in page.frames:
            for sel in selectors:
                try:
                    loc = frame.locator(sel).first
                    if loc.is_visible(timeout=2000):
                        return loc
                except Exception:
                    continue
        return None

    pages_text = {}
    page = None
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(LOGIN_URL, timeout=60000)
            page.wait_for_timeout(6000)  # Lightning login form renders late
            user_box = find_in_frames(page, USER_SELECTORS)
            if not user_box:
                raise Exception("could not find username field on login page")
            user_box.fill(user)
            pass_box = find_in_frames(page, PASS_SELECTORS)
            if not pass_box:
                raise Exception("could not find password field on login page")
            pass_box.fill(pwd)
            btn = find_in_frames(page, SUBMIT_SELECTORS)
            if btn:
                btn.click()
            else:
                pass_box.press("Enter")
            page.wait_for_timeout(8000)
            if "login" in page.url.lower():
                raise Exception("still on login page after submit - check credentials")
            for name, url in PORTAL_PAGES.items():
                page.goto(url, timeout=60000)
                page.wait_for_timeout(6000)  # Lightning components render late
                pages_text[name] = page.inner_text("body")
                harvest_links(page.eval_on_selector_all(
                    "a[href]", "els => els.map(e => ({text: e.textContent, href: e.href}))"
                ), ITEM_LINKS)
                print(f"Phase 1 - pulled {name} ({len(pages_text[name])} chars)")
            # visit each item's detail page - Service/Scope/Outcomes/Product etc.
            # are key fields for the managers' report
            for ref, url in list(ITEM_LINKS.items())[:MAX_DETAIL_PAGES]:
                try:
                    page.goto(url, timeout=60000)
                    page.wait_for_timeout(5000)
                    pages_text[f"detail {ref}"] = page.inner_text("body")[:5000]
                    print(f"Phase 1 - pulled detail for {ref}")
                except Exception as de:
                    print(f"Phase 1 - detail pull failed for {ref}: {de}")
            browser.close()
        return pages_text
    except Exception as e:
        print(f"Phase 1 - live pull failed: {e}")
        # dump what the robot saw so the selectors can be fixed precisely
        # (source/ is gitignored - these never get committed)
        try:
            if page:
                os.makedirs(SOURCE_DIR, exist_ok=True)
                page.screenshot(path=os.path.join(SOURCE_DIR, "_login_debug.png"), full_page=True)
                with open(os.path.join(SOURCE_DIR, "_login_debug.html"), "w", encoding="utf-8") as f:
                    f.write(page.content())
                print("Phase 1 - saved source\\_login_debug.png - send this screenshot to Claude to fix the login")
        except Exception:
            pass
        return None

def read_saved_source():
    """Fallback: parse the newest page Kevin saved into source/ from a
    logged-in browser (Ctrl+S the FlexPoints page as HTML, or paste as .txt)."""
    files = [f for f in glob.glob(os.path.join(SOURCE_DIR, "*"))
             if os.path.isfile(f)
             and not f.lower().endswith((".md", ".gitkeep", ".png"))
             and not os.path.basename(f).startswith("_")]
    if not files:
        return None
    newest = max(files, key=os.path.getmtime)
    raw = open(newest, "r", encoding="utf-8", errors="replace").read()
    if newest.lower().endswith((".html", ".htm", ".mhtml")):
        harvest_links([{"text": re.sub(r"<[^>]+>", "", m.group(2)), "href": m.group(1)}
                       for m in re.finditer(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>',
                                            raw, re.S | re.I)], ITEM_LINKS)
        raw = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", raw, flags=re.S | re.I)
        raw = re.sub(r"<[^>]+>", " ", raw)
        raw = re.sub(r"\s+", " ", raw)
    print(f"Phase 1 - using saved source {os.path.basename(newest)} "
          f"(saved {datetime.fromtimestamp(os.path.getmtime(newest)):%Y-%m-%d %H:%M})")
    return {"saved_page": raw}

print("Phase 1 - pulling FlexPoints pages...")
ITEM_LINKS = {}
pages = get_portal_text() or read_saved_source()
if not pages:
    print("ERROR: no portal data - set ACCESS_PORTAL_USER/ACCESS_PORTAL_PASSWORD env vars,")
    print("       or save the FlexPoints page from a logged-in browser into source/")
    raise SystemExit(1)

# -- Phase 2 - AI parses pages into structured report ------------------------
print("Phase 2 - calling Anthropic API to build report...")

combined = "\n\n".join(f"=== PAGE: {k} ===\n{v[:60000]}" for k, v in pages.items())

PROMPT = """You are building a recurring FlexPoints status report for Kevin Lelitte
(HR Systems, University of Oxford) from pages of the Access Group Customer
Success Portal. FlexPoints are Access service credits: allocated by contract
value, roll over for 12 months, spent on training/consultancy/configuration
services booked by the organisation's Points Guardian.

The FlexPoints page shows: a summary tile row (Available / Requested / Booked /
Expiring within 30, 60, 90 Days / Awaiting Estimate / Awaiting Approval), then
tables: Quote Received, Requested, Booked, and Transaction History (which adds
an Expiry Date column; points top-ups are positive, spends negative). The page
also lists the organisation's FlexPoints Guardians (name + email). Pages named
"detail <ref>" are individual transaction/case detail pages with fields like
Service, Description, Agreed Scope, Outcomes, Prerequisites, Type, Product,
Requester, Assigned to - this detail is key for the managers' report.

From the page text below, return ONLY a JSON object with this shape:
{
  "summary_tiles": {
    "available": <int or null>, "requested": <int or null>, "booked": <int or null>,
    "expiring_30": <int or null>, "expiring_60": <int or null>, "expiring_90": <int or null>,
    "awaiting_estimate": <int or null>, "awaiting_approval": <int or null>
  },
  "quotes_received":     [{"ref": "", "name": "", "owner": "", "date": "", "points": <int>, "status": "", "detail": <see below>}],
  "requested":           [{"ref": "", "name": "", "owner": "", "date": "", "points": <int>, "status": "", "detail": <see below>}],
  "booked":              [{"ref": "", "name": "", "owner": "", "date": "", "points": <int>, "status": "", "detail": <see below>}],
  "transaction_history": [{"ref": "", "name": "", "owner": "", "date": "", "expiry_date": "<or null>", "points": <int>, "status": "", "detail": <see below>}],
  "guardians": [{"name": "", "email": ""}],
  "highlights": ["<up to 5 short bullets: points at risk of expiry, quotes awaiting approval, requested items stuck unbooked, anything Kevin should act on>"],
  "context": "<2-4 sentence plain-English status paragraph, lead with the most urgent action>"
}

"detail" (include ONLY when a "detail <ref>" page exists for that ref, else omit):
  {"service": "", "type": "", "product": "", "agreed_scope": "",
   "description": "<max 30 words>", "outcomes": "<max 30 words>",
   "requester": "", "assigned_to": ""}

Rules: capture EVERY row of every table - this is a full report, not a sample.
Keep points signs exactly as shown (spends negative, top-ups positive). Only
report figures actually present in the page text - use null/empty lists where
the pages do not show a value. Never invent points balances.

PAGE TEXT:
""" + combined

client = anthropic.Anthropic()
resp = client.messages.create(
    model="claude-haiku-4-5",
    max_tokens=4000,
    messages=[{"role": "user", "content": PROMPT}]
)
raw = resp.content[0].text.strip()
raw = re.sub(r"^```(json)?|```$", "", raw, flags=re.M).strip()
parsed = json.loads(raw)

report = {
    "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "source": "live portal" if "flexpoints" in pages else "saved page",
    **parsed
}

# link every item to its exact transaction page on the portal
linked = 0
for key in ("quotes_received", "requested", "booked", "transaction_history"):
    for item in report.get(key) or []:
        url = ITEM_LINKS.get(str(item.get("ref", "")).strip())
        if url:
            item["url"] = url
            linked += 1
print(f"Phase 2 - {linked} item(s) linked to portal detail pages")
print(f"Phase 2 done - quotes:{len(report.get('quotes_received', []))} "
      f"requested:{len(report.get('requested', []))} "
      f"booked:{len(report.get('booked', []))} "
      f"history:{len(report.get('transaction_history', []))} "
      f"available:{report.get('summary_tiles', {}).get('available')}")

# -- Phase 3 - append snapshot to history -------------------------------------

def gh_get(path):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"token {GITHUB_PAT}", "User-Agent": "ag-flexpoints-script"})
    try:
        with urllib.request.urlopen(req) as r:
            body = json.loads(r.read())
            return json.loads(base64.b64decode(body["content"])), body["sha"]
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None, None
        raise

print("Phase 3 - updating history...")
history, history_sha = (gh_get(HISTORY_PATH) if GITHUB_PAT else (None, None))
history = history or {"snapshots": []}
history["snapshots"] = [s for s in history["snapshots"]
                        if s.get("generated", "")[:10] != report["generated"][:10]]
tiles = report.get("summary_tiles", {})
history["snapshots"].append({
    "generated": report["generated"],
    "available": tiles.get("available"),
    "requested": tiles.get("requested"),
    "booked":    tiles.get("booked"),
})
history["snapshots"] = history["snapshots"][-104:]  # ~2 years of weekly runs
print(f"Phase 3 done - {len(history['snapshots'])} snapshot(s) in history")

# -- Phase 4 - push to GitHub --------------------------------------------------
print("Phase 4 - pushing report to GitHub...")

def gh_put(path, obj, sha=None):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
    headers = {"Authorization": f"token {GITHUB_PAT}",
               "Content-Type": "application/json",
               "User-Agent": "ag-flexpoints-script"}
    if sha is None:
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req) as r:
                sha = json.loads(r.read()).get("sha")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
    payload = {
        "message": f"chore: update flexpoints report {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "content": base64.b64encode(
            json.dumps(obj, indent=2, ensure_ascii=False).encode("utf-8")).decode("ascii")
    }
    if sha:
        payload["sha"] = sha
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"),
                                 headers=headers, method="PUT")
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())["commit"]["sha"][:7]

if not GITHUB_PAT:
    print("ERROR: GITHUB_PAT env var not set - cannot push.")
    out = os.path.join(SCRIPT_DIR, "data")
    os.makedirs(out, exist_ok=True)
    json.dump(report, open(os.path.join(out, "flexpoints.json"), "w",
              encoding="utf-8"), indent=2, ensure_ascii=False)
    print(f"Report written locally to {out}\\flexpoints.json instead.")
else:
    c1 = gh_put(REPORT_PATH, report)
    c2 = gh_put(HISTORY_PATH, history, history_sha)
    print(f"Phase 4 done - report pushed (commits {c1}, {c2})")

print("All done.")
