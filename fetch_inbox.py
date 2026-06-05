import json, os, sys
from datetime import datetime, timedelta, timezone
import msal, requests
# -- Config
CLIENT_ID   = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Azure CLI (pre-registered)
AUTHORITY   = "https://login.microsoftonline.com/common"
SCOPES      = ["Mail.Read", "Calendars.Read"]
CACHE_FILE  = os.path.join(os.path.dirname(__file__), "msal_token_cache.json")
OUTPUT_FILE = os.path.join(os.path.dirname(__file__), "data", "inbox_raw.json")
# -- Token cache
cache = msal.SerializableTokenCache()
if os.path.exists(CACHE_FILE):
    cache.deserialize(open(CACHE_FILE).read())
app = msal.PublicClientApplication(
    CLIENT_ID, authority=AUTHORITY, token_cache=cache
)
# -- Auth: silent first, device code fallback
accounts = app.get_accounts()
result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
if not result:
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print("ERROR: Failed to create device flow.", flow)
        sys.exit(1)
    print("\n" + flow["message"] + "\n")
    result = app.acquire_token_by_device_flow(flow)
if "access_token" not in result:
    print("ERROR: Auth failed.", result.get("error_description"))
    sys.exit(1)
# Save refreshed cache
with open(CACHE_FILE, "w") as f:
    f.write(cache.serialize())
token   = result["access_token"]
headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
def graph_get(url):
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.json()
# -- Date ranges
now       = datetime.now(timezone.utc)
seven_ago = (now - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
today_s   = now.strftime("%Y-%m-%dT00:00:00Z")
tomorrow_e = (now + timedelta(days=2)).strftime("%Y-%m-%dT00:00:00Z")
# -- Pull inbox (last 7 days, up to 50)
print("Pulling inbox...")
inbox_url = (
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
    f"?$filter=receivedDateTime ge {seven_ago}"
    "&$orderby=receivedDateTime desc"
    "&$top=50"
    "&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,body,hasAttachments,importance"
)
inbox = graph_get(inbox_url)
# -- Pull sent items (last 7 days, up to 50)
print("Pulling sent items...")
sent_url = (
    "https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages"
    f"?$filter=sentDateTime ge {seven_ago}"
    "&$orderby=sentDateTime desc"
    "&$top=50"
    "&$select=id,subject,toRecipients,sentDateTime,bodyPreview"
)
sent = graph_get(sent_url)
# -- Pull calendar (today + tomorrow)
print("Pulling calendar...")
cal_url = (
    "https://graph.microsoft.com/v1.0/me/calendarView"
    f"?startDateTime={today_s}&endDateTime={tomorrow_e}"
    "&$orderby=start/dateTime"
    "&$select=id,subject,start,end,location,organizer,attendees,bodyPreview,isAllDay"
)
calendar = graph_get(cal_url)
# -- Assemble output
output = {
    "pulled_at": now.isoformat(),
    "inbox":     inbox.get("value", []),
    "sent":      sent.get("value", []),
    "calendar":  calendar.get("value", [])
}
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(output, f, indent=2, ensure_ascii=False)
print(f"\n✅ Done. Output written to {OUTPUT_FILE}")
print(f"   Inbox:    {len(output['inbox'])} messages")
print(f"   Sent:     {len(output['sent'])} messages")
print(f"   Calendar: {len(output['calendar'])} events")
