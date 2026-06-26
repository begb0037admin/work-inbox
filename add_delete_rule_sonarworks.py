"""
add_delete_rule_sonarworks.py
Adds a single Outlook delete rule for sonarworks.com marketing emails.
Run once — idempotent (skips if rule already exists).
"""

import win32com.client


RULE_NAME = "Del - Sonarworks"
SENDER_DOMAIN = "@sonarworks.com"


def main():
    print("Connecting to Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    rules = ns.DefaultStore.GetRules()

    # Check if rule already exists
    for i in range(rules.Count):
        if rules.Item(i + 1).Name == RULE_NAME:
            print(f"Rule '{RULE_NAME}' already exists. Nothing to do.")
            input("Press Enter to close...")
            return

    print(f"Creating rule: {RULE_NAME}...")
    rule = rules.Create(RULE_NAME, 0)  # 0 = olRuleReceive

    # Condition: sender address contains @sonarworks.com
    cond = rule.Conditions.SenderAddress
    cond.Address = [SENDER_DOMAIN]
    cond.Enabled = True

    # Action: delete (move to Deleted Items)
    rule.Actions.Delete.Enabled = True

    rules.Save()
    print(f"Done. Rule '{RULE_NAME}' created — all incoming sonarworks.com emails will go to Deleted Items.")
    input("Press Enter to close...")


if __name__ == "__main__":
    main()
