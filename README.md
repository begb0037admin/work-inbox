# work-inbox

**Live dashboard:** https://begb0037admin.github.io/work-inbox/

Fully automated morning inbox briefing for Kevin Lelitte, HR Systems, University of Oxford.

`fetch_inbox.py` runs on schedule → Anthropic triage → pushes `data/briefing.json` → dashboard renders live.

---

## Quick Load

Paste into Claude chat to load project context:

| File | URL |
|---|---|
| `CLAUDE.md` | https://raw.githubusercontent.com/begb0037admin/work-inbox/main/CLAUDE.md |
| `HANDOVER.md` | https://raw.githubusercontent.com/begb0037admin/work-inbox/main/HANDOVER.md |
