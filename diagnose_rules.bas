' Outlook Rules Diagnostic
' Paste into Outlook VBA editor (Alt+F11 > Insert > Module) then F5
' Tests each step of MoveToFolder rule creation to find exactly where it fails.

Option Explicit

Sub DiagnoseRules()
    Dim ns       As Outlook.NameSpace
    Dim inbox    As Outlook.MAPIFolder
    Dim oRules   As Outlook.Rules
    Dim testRule As Outlook.Rule
    Dim mta      As Object   ' MoveOrCopyRuleAction - use Object to avoid type issues
    Dim rptFld   As Outlook.MAPIFolder
    Dim log      As String
    Dim i        As Integer
    
    log = "Outlook Rules Diagnostic" & vbCrLf & "=========================" & vbCrLf
    
    ' Step 1: Get namespace and inbox
    On Error Resume Next
    Set ns = Application.GetNamespace("MAPI")
    If Err.Number <> 0 Then
        MsgBox "FAIL Step 1 (GetNamespace): " & Err.Description, vbCritical
        Exit Sub
    End If
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    If Err.Number <> 0 Then
        MsgBox "FAIL Step 1 (GetDefaultFolder): " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    log = log & "Step 1 OK: namespace and inbox" & vbCrLf
    
    ' Step 2: Find Reports folder
    On Error Resume Next
    Set rptFld = inbox.Folders("Reports")
    If Err.Number <> 0 Then
        log = log & "Step 2 FAIL: inbox.Folders('Reports') - " & Err.Description & vbCrLf
        Set rptFld = inbox   ' fall back to inbox itself for further tests
    ElseIf rptFld Is Nothing Then
        log = log & "Step 2 FAIL: inbox.Folders('Reports') returned Nothing" & vbCrLf
        Set rptFld = inbox
    Else
        log = log & "Step 2 OK: Reports folder - " & rptFld.FolderPath & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Step 3: Get rules
    On Error Resume Next
    Set oRules = ns.DefaultStore.GetRules()
    If Err.Number <> 0 Then
        log = log & "Step 3 FAIL: GetRules - " & Err.Description & vbCrLf
        MsgBox log, vbCritical, "Diagnostic"
        Exit Sub
    ElseIf oRules Is Nothing Then
        log = log & "Step 3 FAIL: GetRules returned Nothing" & vbCrLf
        MsgBox log, vbCritical, "Diagnostic"
        Exit Sub
    End If
    On Error GoTo 0
    log = log & "Step 3 OK: GetRules - " & oRules.Count & " existing rules" & vbCrLf
    
    ' Remove stale test rule if present
    For i = 1 To oRules.Count
        If oRules.Item(i).Name = "DIAG_TEST" Then
            oRules.Remove i
            oRules.Save
            Set oRules = ns.DefaultStore.GetRules()
            Exit For
        End If
    Next i
    
    ' Step 4: Create a test rule
    On Error Resume Next
    Set testRule = oRules.Create("DIAG_TEST", olRuleReceive)
    If Err.Number <> 0 Then
        log = log & "Step 4 FAIL: Rules.Create - " & Err.Description & vbCrLf
        MsgBox log, vbCritical, "Diagnostic"
        Exit Sub
    ElseIf testRule Is Nothing Then
        log = log & "Step 4 FAIL: Rules.Create returned Nothing" & vbCrLf
        MsgBox log, vbCritical, "Diagnostic"
        Exit Sub
    End If
    On Error GoTo 0
    log = log & "Step 4 OK: Rule created" & vbCrLf
    
    ' Step 5: Access MoveToFolder action
    On Error Resume Next
    Set mta = testRule.Actions.MoveToFolder
    If Err.Number <> 0 Then
        log = log & "Step 5 FAIL: Actions.MoveToFolder - " & Err.Description & vbCrLf
        GoTo Cleanup
    ElseIf mta Is Nothing Then
        log = log & "Step 5 FAIL: Actions.MoveToFolder is Nothing" & vbCrLf
        GoTo Cleanup
    End If
    On Error GoTo 0
    log = log & "Step 5 OK: MoveToFolder action accessible" & vbCrLf
    
    ' Step 6: Set folder to Reports
    On Error Resume Next
    Set mta.Folder = rptFld
    If Err.Number <> 0 Then
        log = log & "Step 6 FAIL: Set MoveToFolder.Folder = Reports - " & Err.Description & " (" & Err.Number & ")" & vbCrLf
        GoTo Cleanup
    End If
    On Error GoTo 0
    log = log & "Step 6 OK: Folder set to " & rptFld.Name & vbCrLf
    
    ' Step 7: Enable and add a sender condition
    On Error Resume Next
    mta.Enabled = True
    testRule.Conditions.SenderAddress.Address = Array("itservxt@ox.ac.uk")
    testRule.Conditions.SenderAddress.Enabled = True
    If Err.Number <> 0 Then
        log = log & "Step 7 FAIL: Set conditions - " & Err.Description & vbCrLf
        GoTo Cleanup
    End If
    On Error GoTo 0
    log = log & "Step 7 OK: Conditions set" & vbCrLf
    
    ' Step 8: Save
    On Error Resume Next
    oRules.Save
    If Err.Number <> 0 Then
        log = log & "Step 8 FAIL: Save - " & Err.Description & vbCrLf
        GoTo Cleanup
    End If
    On Error GoTo 0
    log = log & "Step 8 OK: Saved" & vbCrLf
    
Cleanup:
    Err.Clear
    On Error Resume Next
    Set oRules = ns.DefaultStore.GetRules()
    For i = 1 To oRules.Count
        If oRules.Item(i).Name = "DIAG_TEST" Then
            oRules.Remove i
            oRules.Save
            log = log & vbCrLf & "Test rule removed."
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    MsgBox log, vbInformation, "Diagnostic"
End Sub
