' Outlook Inbox File Rules Setup
' Run from inside Outlook: Alt+F11 > Insert > Module > paste this > F5
' Creates 13 auto-file rules. Safe to re-run - skips rules that already exist.

Option Explicit

Dim gCreated As Integer
Dim gSkipped As Integer
Dim gFailed  As Integer
Dim gLog     As String

Sub CreateInboxFileRules()
    Dim ns    As Outlook.NameSpace
    Dim inbox As Outlook.MAPIFolder

    Set ns    = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    gCreated = 0: gSkipped = 0: gFailed = 0: gLog = ""

    FR ns, inbox, "File - Reports - ITSRVXT", _
       inbox.Folders("Reports"), _
       addr:=Array("itservxt@ox.ac.uk")

    FR ns, inbox, "File - Reports - PeopleXD Reports", _
       inbox.Folders("Reports"), _
       addr:=Array("PeopleXDReports@theaccessgroup.com")

    FR ns, inbox, "File - Access Support - Team Cases", _
       inbox.Folders("Access Group").Folders("Team Cases"), _
       addr:=Array("support.access@theaccessgroup.com"), useCC:=True

    FR ns, inbox, "File - Access Support - My Cases", _
       inbox.Folders("Access Group").Folders("My Cases"), _
       addr:=Array("support.access@theaccessgroup.com")

    FR ns, inbox, "File - PeopleXD System", _
       inbox.Folders("PeopleXD System"), _
       addr:=Array("peoplexd@accessacloud.com")

    FR ns, inbox, "File - Cority", _
       inbox.Folders("H&S").Folders("Cority"), _
       hdr:=Array("@cority.com")

    FR ns, inbox, "File - HR Broadcast", _
       inbox.Folders("Reference").Folders("HR Broadcast"), _
       addr:=Array("hris@admin.ox.ac.uk")

    FR ns, inbox, "File - ICT subject tag", _
       inbox.Folders("Reference").Folders("ICT Mailing Lists"), _
       subj:=Array("[ict-a]")

    FR ns, inbox, "File - ICT senders", _
       inbox.Folders("Reference").Folders("ICT Mailing Lists"), _
       addr:=Array("changenotifications@it.ox.ac.uk", "skills@it.ox.ac.uk")

    FR ns, inbox, "File - Bodleian & Sector", _
       inbox.Folders("Reference").Folders("Bodleian & Sector"), _
       hdr:=Array("@bodleian.ox.ac.uk", "@jiscmail.ac.uk")

    FR ns, inbox, "File - Team James", _
       inbox.Folders("Team").Folders("James Salas Guillen"), _
       addr:=Array("james.salas-guillen@admin.ox.ac.uk")

    FR ns, inbox, "File - Team Michael", _
       inbox.Folders("Team").Folders("Michael O'Sullivan"), _
       addr:=Array("michael.osullivan@admin.ox.ac.uk")

    FR ns, inbox, "File - Team Asta", _
       inbox.Folders("Team").Folders("Asta Palmer"), _
       addr:=Array("asta.palmer@admin.ox.ac.uk")

    MsgBox "Done" & vbCrLf & _
           "Created : " & gCreated & vbCrLf & _
           "Skipped : " & gSkipped & vbCrLf & _
           "Failed  : " & gFailed  & vbCrLf & vbCrLf & gLog, _
           vbInformation, "Inbox Rules Setup"
End Sub

' ---------------------------------------------------------------------------
Sub FR(ns As Outlook.NameSpace, inbox As Outlook.MAPIFolder, _
        ruleName As String, destFolder As Outlook.MAPIFolder, _
        Optional addr  As Variant, _
        Optional subj  As Variant, _
        Optional hdr   As Variant, _
        Optional useCC As Boolean = False)

    Dim oRules As Outlook.Rules
    Set oRules = ns.DefaultStore.GetRules()

    ' Skip if already exists
    Dim i As Integer
    For i = 1 To oRules.Count
        If oRules.Item(i).Name = ruleName Then
            gSkipped = gSkipped + 1
            gLog = gLog & "  skip    : " & ruleName & vbCrLf
            Exit Sub
        End If
    Next i

    On Error GoTo Fail
    Dim r As Outlook.Rule
    Set r = oRules.Create(ruleName, olRuleReceive)

    If Not IsMissing(addr) Then
        r.Conditions.SenderAddress.Address = addr
        r.Conditions.SenderAddress.Enabled = True
    End If
    If Not IsMissing(subj) Then
        r.Conditions.Subject.Text = subj
        r.Conditions.Subject.Enabled = True
    End If
    If Not IsMissing(hdr) Then
        r.Conditions.MessageHeader.Text = hdr
        r.Conditions.MessageHeader.Enabled = True
    End If
    If useCC Then
        r.Conditions.CC.Enabled = True
    End If

    Set r.Actions.MoveToFolder.Folder = destFolder
    r.Actions.MoveToFolder.Enabled = True
    oRules.Save

    gCreated = gCreated + 1
    gLog = gLog & "  created : " & ruleName & vbCrLf
    Exit Sub

Fail:
    gFailed = gFailed + 1
    gLog = gLog & "  FAIL    : " & ruleName & " - " & Err.Description & vbCrLf
    On Error GoTo 0
End Sub
