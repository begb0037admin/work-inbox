"""Self-test for the Command Centre triage identity fix (Phase 3.5/3.6).

Extracts the real `_cc_mail_key` / `_owa_link` helpers from fetch_inbox.py by
AST (no import side effects -- fetch_inbox.py is a top-to-bottom script with no
__main__ guard) and asserts:

  1. COM parity  -- when entry_id is a real Outlook EntryID, the dedup/ledger
     key is byte-identical to the pre-fix behaviour (entry_id + "_" + task_id),
     so nothing changes for the Outlook COM pipeline.
  2. IMAP fix    -- when entry_id == "" (every IMAP message), the key falls
     back to the internet Message-ID (angle brackets stripped) instead of the
     old poisoned "" -> "_<task_id>" collision.
  3. Skip guard  -- key is "" only when BOTH entry_id and message_id are
     missing; the promote/ledger code must skip such a candidate.
  4. OWA link    -- built only from a non-empty Message-ID, on the
     outlook.office.com host the dashboard allowlist accepts.

Run:  python test_cc_mail_key.py      (exit 0 = pass)
"""
import ast
import os
import sys
import urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "fetch_inbox.py")


def _load_helpers():
    tree = ast.parse(open(SRC, encoding="utf-8").read(), filename=SRC)
    want = {"_cc_mail_key", "_owa_link"}
    ns = {"urllib": urllib}
    found = set()
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in want:
            code = compile(ast.Module(body=[node], type_ignores=[]), SRC, "exec")
            exec(code, ns)
            found.add(node.name)
    missing = want - found
    if missing:
        raise SystemExit("FAIL: helpers not found in fetch_inbox.py: %s" % missing)
    return ns["_cc_mail_key"], ns["_owa_link"]


def main():
    mail_key, owa_link = _load_helpers()
    fails = []

    def check(name, cond):
        print(("  ok   " if cond else "  FAIL ") + name)
        if not cond:
            fails.append(name)

    REAL_EID = "0000000060196AC9D4535F45A195B2716E93E76B0700FA1BE8"  # COM EntryID shape
    MID = "<CY8PR0102MB1234ABCDEF@CY8PR0102MB1234.prod.outlook.com>"
    MID_BARE = "CY8PR0102MB1234ABCDEF@CY8PR0102MB1234.prod.outlook.com"
    TID = "t2608191643001"

    # 1. COM parity: key identical to the historical `entry_id + "_" + task_id`
    check("COM: mail_key == raw EntryID",
          mail_key(REAL_EID, "") == REAL_EID)
    check("COM: mail_key unaffected by an also-present message_id",
          mail_key(REAL_EID, MID) == REAL_EID)
    check("COM: applied-ledger key byte-identical to pre-fix",
          mail_key(REAL_EID, "") + "_" + TID == REAL_EID + "_" + TID)

    # 2. IMAP fix: entry_id "" -> fall back to Message-ID (brackets stripped)
    check("IMAP: entry_id '' falls back to bare Message-ID",
          mail_key("", MID) == MID_BARE)
    check("IMAP: key is NOT the poisoned '_<task_id>' form",
          mail_key("", MID) + "_" + TID == MID_BARE + "_" + TID
          and (mail_key("", MID) + "_" + TID) != "_" + TID)
    check("IMAP: two different messages produce different keys",
          mail_key("", "<a@x>") != mail_key("", "<b@x>"))
    check("IMAP: same message, bracketed vs bare, produces the same key",
          mail_key("", MID) == mail_key("", MID_BARE))

    # 3. Skip guard: empty only when BOTH are missing
    check("no identity: both missing -> '' (caller must skip)",
          mail_key("", "") == "" and mail_key(None, None) == "")

    # 4. OWA link
    link = owa_link(MID)
    check("OWA link built from Message-ID, https + outlook.office.com host",
          link.startswith("https://outlook.office.com/mail/search?query=")
          and urllib.parse.quote(MID_BARE) in link)
    check("OWA link empty for empty Message-ID",
          owa_link("") == "" and owa_link(None) == "")

    print("")
    if fails:
        print("RESULT: %d FAILED" % len(fails))
        return 1
    print("RESULT: all passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
