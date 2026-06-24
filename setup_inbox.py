"""
Outlook Inbox Setup v3
Creates folder structure and rules per docs/INBOX_ORGANISATION.md
Governance: pull from GitHub, run via Setup_Inbox.bat
"""
import win32com.client

VERSION = "v3-incremental"

def main():
    print(f"Outlook Inbox Setup {VERSION}")
    print("===================")

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns      = outlook.GetNamespace("MAPI")
    inbox   = ns.GetDefaultFolder(6)   # olFolderInbox

    # -- Phase 1: Folder structure --
    print("\nPhase 1 - Creating folders...")

    def mkf(parent, name):
        try:
            return parent.Folders[name]
        except Exception:
            f = parent.Folders.Add(name)
            print(f"  + {parent.Name} / {name}")
            return f

    ag   = mkf(inbox, "Access Group")
    mkf(ag, "My Cases")
    mkf(ag, "Team Cases")
    mkf(inbox, "PeopleXD System")
    mkf(inbox, "Reports")
    team = mkf(inbox, "Team")
    mkf(team, "Michael O'Sullivan")
    mkf(team, "Asta Palmer")
    mkf(team, "James Salas Guillen")
    hs   = mkf(inbox, "H&S")
    mkf(hs, "Cority")
    mkf(hs, "DSE . IRIS . Risk Base")
    proj = mkf(inbox, "Projects")
    mkf(proj, "DTP1092 - Colleges & REF2029")
    mkf(proj, "DTP1334 - H&S System Evaluation")
    mkf(proj, "ePloy")
    ref  = mkf(inbox, "Reference")
    mkf(ref, "HR Broadcast")
    mkf(ref, "ICT Mailing Lists")
    mkf(ref, "Bodleian & Sector")
    mkf(inbox, "_Archive")

    print("Phase 1 done")

    def fld(*path):
        f = inbox
        for p in path:
            f = f.Folders[p]
        return f

    # -- Phase 2: Rules (incremental saves to identify failures) --
    print("\nPhase 2 - Creating rules...")

    rules_obj = ns.DefaultStore.GetRules()

    existing = set()
    for i in range(1, rules_obj.Count + 1):
        try:
            existing.add(rules_obj.Item(i).Name)
        except Exception:
            pass

    counts = {"created": 0, "skipped": 0, "failed": 0}

    def save_one(name):
        """Try to save. If rejected, remove this rule and restore clean state."""
        try:
            rules_obj.Save()
            counts["created"] += 1
            print(f"  created: {name}")
        except Exception as se:
            for i in range(rules_obj.Count, 0, -1):
                try:
                    if rules_obj.Item(i).Name == name:
                        rules_obj.Remove(i)
                        try:
                            rules_obj.Save()
                        except Exception:
                            pass
                        break
                except Exception:
                    pass
            counts["failed"] += 1
            print(f"  FAIL (Outlook rejected): {name}")

    def rule_delete(name, addresses=None, subjects=None, headers=None):
        if name in existing:
            counts["skipped"] += 1
            print(f"  skip (exists): {name}")
            return
        try:
            r = rules_obj.Create(name, 0)
            if addresses:
                r.Conditions.SenderAddress.Address = addresses
                r.Conditions.SenderAddress.Enabled = True
            if subjects:
                r.Conditions.Subject.Text = subjects
                r.Conditions.Subject.Enabled = True
            if headers:
                r.Conditions.MessageHeader.Text = headers
                r.Conditions.MessageHeader.Enabled = True
            r.Actions.Delete.Enabled = True
            save_one(name)
        except Exception as e:
            counts["failed"] += 1
            print(f"  FAIL (create error): {name} - {e}")

    def rule_file(name, dest, addresses=None, subjects=None, headers=None, cc=False):
        if name in existing:
            counts["skipped"] += 1
            print(f"  skip (exists): {name}")
            return
        try:
            r = rules_obj.Create(name, 0)
            if addresses:
                r.Conditions.SenderAddress.Address = addresses
                r.Conditions.SenderAddress.Enabled = True
            if subjects:
                r.Conditions.Subject.Text = subjects
                r.Conditions.Subject.Enabled = True
            if headers:
                r.Conditions.MessageHeader.Text = headers
                r.Conditions.MessageHeader.Enabled = True
            if cc:
                r.Conditions.CC.Enabled = True
            r.Actions.MoveToFolder.Folder = dest
            r.Actions.MoveToFolder.Enabled = True
            save_one(name)
        except Exception as e:
            counts["failed"] += 1
            print(f"  FAIL (create error): {name} - {e}")

    # Auto-delete rules
    rule_delete("Del - Teams notifications",      addresses=["no-reply@teams.mail.microsoft"])
    rule_delete("Del - GitHub",                   addresses=["notifications@github.com", "noreply@github.com"])
    rule_delete("Del - Access Group marketing",   headers=["@go.theaccessgroup.com", "@surveys.theaccessgroup.com"])
    rule_delete("Del - New Vacancy Notification", subjects=["New Vacancy Notification"])
    rule_delete("Del - Cority status alerts",     addresses=["csn@mail.status.cority.com"])
    rule_delete("Del - DistroKid",                headers=["@hello.distrokid.com"])
    rule_delete("Del - Anthropic/Claude",         addresses=["no-reply@email.claude.com"])
    rule_delete("Del - Descript marketing",       headers=["@marketing.descript.com"])
    rule_delete("Del - Accessplanit",             headers=["@accessplanit.com"])
    rule_delete("Del - Skype voicemail",          addresses=["no-reply@emails.skype.com"])
    rule_delete("Del - MetaCompliance",           headers=["@metacompliance.com"])
    rule_delete("Del - Annual Leave system",      subjects=["ANNUAL LEAVE request submitted"])

    # Auto-file rules
    rule_file("File - Reports - ITSRVXT",           dest=fld("Reports"),                         addresses=["itservxt@ox.ac.uk"])
    rule_file("File - Reports - PeopleXD Reports",  dest=fld("Reports"),                         addresses=["PeopleXDReports@theaccessgroup.com"])
    rule_file("File - Access Support - Team Cases", dest=fld("Access Group", "Team Cases"),      addresses=["support.access@theaccessgroup.com"], cc=True)
    rule_file("File - Access Support - My Cases",   dest=fld("Access Group", "My Cases"),        addresses=["support.access@theaccessgroup.com"])
    rule_file("File - PeopleXD System",             dest=fld("PeopleXD System"),                 addresses=["peoplexd@accessacloud.com"])
    rule_file("File - Cority",                      dest=fld("H&S", "Cority"),                   headers=["@cority.com"])
    rule_file("File - HR Broadcast",                dest=fld("Reference", "HR Broadcast"),       addresses=["hris@admin.ox.ac.uk"])
    rule_file("File - ICT subject tag",             dest=fld("Reference", "ICT Mailing Lists"),  subjects=["[ict-a]"])
    rule_file("File - ICT senders",                 dest=fld("Reference", "ICT Mailing Lists"),  addresses=["changenotifications@it.ox.ac.uk", "skills@it.ox.ac.uk"])
    rule_file("File - Bodleian & Sector",           dest=fld("Reference", "Bodleian & Sector"),  headers=["@bodleian.ox.ac.uk", "@jiscmail.ac.uk"])
    rule_file("File - Team James",                  dest=fld("Team", "James Salas Guillen"),      addresses=["james.salas-guillen@admin.ox.ac.uk"])
    rule_file("File - Team Michael",                dest=fld("Team", "Michael O'Sullivan"),       addresses=["michael.osullivan@admin.ox.ac.uk"])
    rule_file("File - Team Asta",                   dest=fld("Team", "Asta Palmer"),             addresses=["asta.palmer@admin.ox.ac.uk"])

    c = counts
    print(f"\nPhase 2 done - created:{c['created']}  skipped:{c['skipped']}  failed:{c['failed']}")
    print("\nSetup complete.")
    if c['failed']:
        print(f"  {c['failed']} rule(s) rejected - see FAIL lines above for which ones.")
        print("  Add those rules manually via Outlook Rules and Alerts.")
    print("  To apply rules to existing emails:")
    print("    Outlook > Home > Rules > Manage Rules & Alerts > Run Rules Now")

if __name__ == "__main__":
    main()
