# COWORK BRIEF — Outlook Inbox Rules Setup
> Execution Seat brief. Follow exactly. No decisions. Report everything back.

**Date:** 24 June 2026
**Issued by:** Claude Code (Reasoning Seat)
**Operator:** Kevin Lelitte, HR Systems Manager/Director, University of Oxford
**Machine:** Kevin's Windows admin desktop (kevin.lelitte@admin.ox.ac.uk)

---

## 1. Who You Are and What This Is

You are the Execution Seat in Kevin's AI operating model (CONSTITUTION.md). Your role is to carry out this brief exactly as written. You do not make decisions. If you encounter something not covered in this brief, you stop and report back — you do not improvise.

This brief is for a one-off Outlook inbox setup task. It will take approximately 20–30 minutes.

---

## 2. Background and Context

Kevin's Outlook inbox (work account: kevin.lelitte@admin.ox.ac.uk, hosted on Oxford University Exchange) has been partially configured by Claude Code during a setup session on 24 June 2026. The work was based on analysis of 2,742 inbox emails and 377 sent emails covering December 2025 to June 2026.

The goal is a clean, self-managing inbox with:
- Noise auto-deleted (unwanted notifications, marketing)
- Known senders auto-filed into organised subfolders
- Only emails needing Kevin's attention staying in the main Inbox

The reference document for the full design is: `begb0037admin/work-inbox` → `docs/INBOX_ORGANISATION.md`

---

## 3. Current Status — What Is Done and What Is Not

| Item | Status | Notes |
|---|---|---|
| Folder structure inside Inbox | ✅ Done | All 19 folders already created |
| 12 auto-delete rules | ✅ Done | Created by script, confirmed in Rules & Alerts |
| 13 auto-file rules (MoveToFolder) | ❌ NOT done | Script failed — must be created manually via UI |
| Existing emails filed to folders | ❌ NOT done | Will happen via Run Rules Now after rules are created |

**Why are the 13 rules not done?** Claude Code attempted to create them via Python COM automation and VBA macro — both failed with Outlook rejecting the MoveToFolder action when created programmatically. They must be created through Outlook's Rules Wizard UI.

**The 12 delete rules are already there.** Do not recreate them. Do not touch them.

---

## 4. What Already Exists — Do Not Touch

The following rules already exist in Outlook Rules & Alerts. Leave them exactly as they are:

**12 rules created by Claude Code (Del- prefix):**
- Del - Teams notifications
- Del - GitHub
- Del - Access Group marketing
- Del - New Vacancy Notification
- Del - Cority status alerts
- Del - DistroKid
- Del - Anthropic/Claude
- Del - Descript marketing
- Del - Accessplanit
- Del - Skype voicemail
- Del - MetaCompliance
- Del - Annual Leave system

**Pre-existing rules Kevin already had (leave completely alone):**
- Accepted: Advanced Insight Reporting...
- Move all messages from Sonarworks Accounts to Sonarworks
- Move all messages from The Sonarworks Team to Sonarworks
- Move all messages from DistroKid to DistroKid
- Move all messages from CoSy - Oxford University to CoSy
- Move all messages from tasknotifications@it.ox.ac.uk to Task Not...
- Move all messages to ITSS Discuss

Total: these 19 rules must remain untouched.

---

## 5. The Folder Structure (already exists — do not create)

All folders are already inside Kevin's Inbox. The structure is:

```
Inbox
├── Access Group
│   ├── My Cases
│   └── Team Cases
├── PeopleXD System
├── Reports
├── Team
│   ├── Michael O'Sullivan
│   ├── Asta Palmer
│   └── James Salas Guillen
├── H&S
│   ├── Cority
│   └── DSE . IRIS . Risk Base
├── Projects
│   ├── DTP1092 - Colleges & REF2029
│   ├── DTP1334 - H&S System Evaluation
│   └── ePloy
├── Reference
│   ├── HR Broadcast
│   ├── ICT Mailing Lists
│   └── Bodleian & Sector
└── _Archive
```

---

## 6. Your Task — 2 Parts

### Part A: Create 13 auto-file rules in Outlook Rules Wizard
### Part B: Run Rules Now to file existing inbox emails

---

## 7. Hard Rules — Read Before Starting

1. **Use Outlook Classic only.** Not the new Outlook. Classic Outlook must be open. If new Outlook is showing, switch it off and open classic Outlook.
2. **Do not modify any existing rules.** Only create new ones.
3. **Do not delete any existing folders or emails.** This is setup only.
4. **Create rules in the exact order listed below.** Order matters — particularly rules 12 and 13 (Access Support) which must be in the right sequence.
5. **If a step fails or produces an unexpected dialog, stop and report back exactly what you see.** Do not guess or work around it.
6. **Screenshot after Part B (Run Rules Now) showing the Inbox email count.** Kevin needs to see the result.

---

## 8. Part A — Creating the 13 Rules

### How to open the Rules Wizard
1. Open Outlook Classic
2. Click the **Home** tab in the ribbon
3. Click **Rules** → **Manage Rules & Alerts**
4. Click **New Rule**
5. Under "Start from a blank rule", select **Apply rule on messages I receive**
6. Click **Next**

You will now be in the Rules Wizard. The steps below describe what to do in the wizard for each rule.

---

### Rule 1 — File - Team James

**Step 1 — Condition:** Tick **"from people or public group"**. Click the underlined link "people or public group" and type: `james.salas-guillen@admin.ox.ac.uk` → click OK.

**Step 2 — Action:** Click Next. Tick **"move it to the specified folder"**. Click the underlined "specified" link. Navigate to **Inbox → Team → James Salas Guillen**. Click OK.

**Step 3 — Finish:** Click Next → Next → Name the rule: `File - Team James` → click Finish.

---

### Rule 2 — File - Team Michael

**Condition:** from `michael.osullivan@admin.ox.ac.uk`
**Action:** Move to **Inbox → Team → Michael O'Sullivan**
**Rule name:** `File - Team Michael`

---

### Rule 3 — File - Team Asta

**Condition:** from `asta.palmer@admin.ox.ac.uk`
**Action:** Move to **Inbox → Team → Asta Palmer**
**Rule name:** `File - Team Asta`

---

### Rule 4 — File - Reports - ITSRVXT

**Condition:** from `itservxt@ox.ac.uk`
**Action:** Move to **Inbox → Reports**
**Rule name:** `File - Reports - ITSRVXT`

---

### Rule 5 — File - Reports - PeopleXD Reports

**Condition:** from `PeopleXDReports@theaccessgroup.com`
**Action:** Move to **Inbox → Reports**
**Rule name:** `File - Reports - PeopleXD Reports`

---

### Rule 6 — File - PeopleXD System

**Condition:** from `peoplexd@accessacloud.com`
**Action:** Move to **Inbox → PeopleXD System**
**Rule name:** `File - PeopleXD System`

---

### Rule 7 — File - Cority

**Condition:** Tick **"with specific words in the message header"**. Click the underlined "specific words" link. Type `@cority.com` → click Add → click OK.
**Action:** Move to **Inbox → H&S → Cority**
**Rule name:** `File - Cority`

---

### Rule 8 — File - HR Broadcast

**Condition:** from `hris@admin.ox.ac.uk`
**Action:** Move to **Inbox → Reference → HR Broadcast**
**Rule name:** `File - HR Broadcast`

---

### Rule 9 — File - ICT subject tag

**Condition:** Tick **"with specific words in the subject"**. Click the underlined link. Type `[ict-a]` → click Add → click OK.
**Action:** Move to **Inbox → Reference → ICT Mailing Lists**
**Rule name:** `File - ICT subject tag`

---

### Rule 10 — File - ICT senders

**Condition:** from people or public group. Add both addresses:
- `changenotifications@it.ox.ac.uk`
- `skills@it.ox.ac.uk`

(In the address box, add the first, click Add, then type the second and click Add again.)

**Action:** Move to **Inbox → Reference → ICT Mailing Lists**
**Rule name:** `File - ICT senders`

---

### Rule 11 — File - Bodleian & Sector

**Condition:** Tick **"with specific words in the message header"**. Click the underlined link. Add both:
- `@bodleian.ox.ac.uk` → Add
- `@jiscmail.ac.uk` → Add
Click OK.

**Action:** Move to **Inbox → Reference → Bodleian & Sector**
**Rule name:** `File - Bodleian & Sector`

---

### Rule 12 — File - Access Support - My Cases

**Condition:** from `support.access@theaccessgroup.com`
**Action:** Move to **Inbox → Access Group → My Cases**
**Rule name:** `File - Access Support - My Cases`

---

### Rule 13 — File - Access Support - Team Cases

This rule has TWO conditions (both must be true):

**Condition 1:** from `support.access@theaccessgroup.com`
**Condition 2:** Tick **"where my name is in the CC box"**

**Action:** Move to **Inbox → Access Group → Team Cases**
**Rule name:** `File - Access Support - Team Cases`

**IMPORTANT:** After saving this rule, it must appear ABOVE Rule 12 (File - Access Support - My Cases) in the rules list. If it is below Rule 12, drag it above. This ensures emails where Kevin is CC'd go to Team Cases, not My Cases.

---

## 9. Part B — Run Rules Now (filing existing emails)

Once all 13 rules are created:

1. In **Manage Rules & Alerts**, click **Run Rules Now**
2. Click **Select All** to select every rule
3. Confirm **Run in Folder: Inbox** is selected
4. **Include subfolders:** leave unticked (Inbox only)
5. Click **Run Now**
6. Wait for completion — this may take several minutes with 365+ emails
7. Click **Close**

**Take a screenshot of the Inbox showing the email count after Run Rules Now completes.**

---

## 10. Verification — What to Check

After Run Rules Now completes, check these folders have emails in them:

- Inbox / Reports (should have ITSRVXT and PeopleXD report emails)
- Inbox / Team / James Salas Guillen
- Inbox / Team / Michael O'Sullivan
- Inbox / Team / Asta Palmer
- Inbox / Reference / HR Broadcast
- Inbox / Reference / ICT Mailing Lists
- Inbox / Access Group / My Cases

Also check Deleted Items — auto-delete rules should have moved Teams notifications, GitHub notifications, etc. there.

---

## 11. What to Report Back

When done, report:

1. Which of the 13 rules were created successfully (list them)
2. Any rules that failed — what exactly happened
3. The Inbox email count before and after Run Rules Now
4. Screenshot of Inbox post-run
5. Any unexpected dialogs, errors, or warnings that appeared during the process
6. Whether Rule 13 (Team Cases) is correctly positioned above Rule 12 (My Cases)

---

## 12. If Something Goes Wrong

- **Rule refuses to save:** Screenshot the error and the rule configuration. Do not try to work around it — report back.
- **Folder not found when navigating:** The folder should already exist. If it is missing, report which folder and stop — do not create folders yourself.
- **Run Rules Now moves something obviously wrong:** Do not try to fix it. Report what happened and where the emails went. Kevin will correct manually.
- **Any other unexpected situation:** Stop. Report exactly what you see. Do not improvise.

---

## Source Documents

- `begb0037admin/work-inbox` → `docs/INBOX_ORGANISATION.md` — full inbox design reference
- `begb0037admin/work-inbox` → `HANDOVER.md` — project state
- `begb0037admin/work-inbox` → `AGENT_MODEL.md` — operating model
- `begb0037admin/work-inbox` → `CONSTITUTION.md` — governance
