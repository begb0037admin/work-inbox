"""
add_delete_rule.py
Adds a single Outlook delete rule for a sender domain or address.
Reusable — run whenever a new marketing sender needs blocking.

Usage:
  python add_delete_rule.py "Del - EcoOnline" "@ecoonline.com"
  python add_delete_rule.py "Del - Example" "noreply@example.com"

Idempotent — skips if the rule name already exists.
"""

import sys
import win32com.client


def main():
    if len(sys.argv) != 3:
        print("Usage: python add_delete_rule.py \"<rule name>\" \"<@domain.com or full address>\"")
        print("Example: python add_delete_rule.py \"Del - EcoOnline\" \"@ecoonline.com\"")
        input("Press Enter to close...")
        return

    rule_name = sys.argv[1]
    sender = sys.argv[2]

    print(f"Connecting to Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    rules = ns.DefaultStore.GetRules()

    for i in range(rules.Count):
        if rules.Item(i + 1).Name == rule_name:
            print(f"Rule '{rule_name}' already exists. Nothing to do.")
            input("Press Enter to close...")
            return

    print(f"Creating rule '{rule_name}' for sender: {sender}")
    rule = rules.Create(rule_name, 0)  # 0 = olRuleReceive

    cond = rule.Conditions.SenderAddress
    cond.Address = [sender]
    cond.Enabled = True

    rule.Actions.Delete.Enabled = True
    rules.Save()

    print(f"Done. All incoming mail from '{sender}' will go to Deleted Items.")
    input("Press Enter to close...")


if __name__ == "__main__":
    main()
