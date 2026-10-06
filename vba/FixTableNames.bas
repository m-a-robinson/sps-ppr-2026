Attribute VB_Name = "modFixTableNames"
Option Explicit

' Rename tables on module sheets (4001SPS, 4NX1SPS ...) to match the sheet:
' _4NX1SPS_Meta, _4NX1SPS_LO and so on. Works on the ACTIVE workbook.
'
' 1. Run FixTableNames with DRY_RUN = True. Nothing changes; the plan is shown
'    in a message box and in full in the Immediate window (Ctrl+G).
' 2. Set DRY_RUN = False and run it again to apply.
'
' Tables already named correctly are left alone, and no rename is made if the
' new name already exists in the workbook. Protected sheets are unprotected
' for the rename and protected again with their original settings.

Private Const DRY_RUN As Boolean = True
Private Const SHEET_PASSWORD As String = ""   ' sheet protection password, if any

Public Sub FixTableNames()
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, nm As Name
    Dim taken As Object, leaving As Object, claimed As Object, prot As Object, toChange As Object
    Dim notes As New Collection, plan As New Collection, temps As New Collection
    Dim items As Variant, p As Variant, sh As Variant, msg As Variant
    Dim sheetName As String, nameNow As String, body As String, suffix As String
    Dim itm As String, target As String, key As String, tempName As String
    Dim report As String, failed As String, errText As String
    Dim i As Long, cut As Long, correct As Long, changed As Boolean

    Set wb = ActiveWorkbook
    items = Array("Meta", "LO", "Assess", "Weekly", "Hours", "Mapping", "Aims", "Syllabus", "Overview", "Notes")

    ' Every name already in use: table names and workbook-level defined names.
    Set taken = CreateObject("Scripting.Dictionary")
    taken.CompareMode = vbTextCompare
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            taken(lo.Name) = True
        Next lo
    Next ws
    For Each nm In wb.Names
        If InStr(nm.Name, "!") = 0 Then taken(nm.Name) = True
    Next nm

    ' Work out what each table on a module sheet should be called.
    For Each ws In wb.Worksheets
        sheetName = ws.Name
        If Not IsModuleSheet(sheetName) Then
            If InStr(1, sheetName, "SPS", vbTextCompare) > 0 And ws.ListObjects.Count > 0 Then
                notes.Add "Sheet """ & sheetName & """ isn't in the 4001SPS format, so its tables were not checked."
            End If
        Else
            For Each lo In ws.ListObjects
                nameNow = lo.Name
                body = nameNow
                If Left$(body, 1) = "_" Then body = Mid$(body, 2)
                cut = InStrRev(body, "_")
                suffix = ""
                If cut > 0 Then suffix = Mid$(body, cut + 1)
                ' "Meta2" or "LO4" come from copied tables: drop the trailing number.
                Do While Len(suffix) > 0 And Right$(suffix, 1) Like "#"
                    suffix = Left$(suffix, Len(suffix) - 1)
                Loop
                itm = ""
                For i = LBound(items) To UBound(items)
                    If StrComp(items(i), suffix, vbTextCompare) = 0 Then
                        itm = items(i)
                        Exit For
                    End If
                Next i
                If itm = "" Then
                    notes.Add sheetName & ": """ & nameNow & """ isn't a proforma table, left alone."
                Else
                    If itm = "Meta" Then CheckMetaCode lo, sheetName, notes
                    target = "_" & sheetName & "_" & itm
                    If StrComp(nameNow, target, vbBinaryCompare) = 0 Then
                        correct = correct + 1
                    Else
                        plan.Add Array(sheetName, nameNow, target)
                    End If
                End If
            Next lo
        End If
    Next ws

    ' Drop any rename whose new name is already in use or already claimed.
    ' Repeat until stable, because a dropped rename keeps its old name.
    changed = True
    Do While changed
        changed = False
        Set leaving = CreateObject("Scripting.Dictionary")
        leaving.CompareMode = vbTextCompare
        Set claimed = CreateObject("Scripting.Dictionary")
        claimed.CompareMode = vbTextCompare
        For i = 1 To plan.Count
            p = plan(i)
            leaving(p(1)) = True
        Next i
        For i = 1 To plan.Count
            p = plan(i)
            key = p(2)
            If (taken.Exists(key) And Not leaving.Exists(key)) Or claimed.Exists(key) Then
                notes.Add p(0) & ": """ & p(1) & """ NOT renamed, because """ & p(2) & """ already exists. Check which table holds the right data."
                plan.Remove i
                changed = True
                Exit For
            End If
            claimed(key) = True
        Next i
    Loop

    For Each msg In notes
        Debug.Print msg
        report = report & msg & vbLf
    Next msg
    For i = 1 To plan.Count
        p = plan(i)
        msg = p(0) & ": """ & p(1) & """ -> """ & p(2) & """"
        Debug.Print msg
        report = report & msg & vbLf
    Next i

    If plan.Count = 0 Then
        Finish report, "Nothing to rename. " & correct & " tables already match their sheet."
        Exit Sub
    End If
    If DRY_RUN Then
        Finish report, "DRY RUN: " & plan.Count & " tables would be renamed, " & correct & " already match." & vbLf & _
                       "Set DRY_RUN = False and run again to apply."
        Exit Sub
    End If

    ' Unprotect only the sheets that need a change, remembering their settings.
    Set prot = CreateObject("Scripting.Dictionary")
    Set toChange = CreateObject("Scripting.Dictionary")
    For i = 1 To plan.Count
        p = plan(i)
        toChange(p(0)) = True
    Next i
    For Each sh In toChange.Keys
        Set ws = wb.Worksheets(sh)
        If IsProtected(ws) Then
            prot(sh) = SaveProtection(ws)
            On Error Resume Next
            ws.Unprotect SHEET_PASSWORD
            On Error GoTo 0
            If IsProtected(ws) Then
                failed = failed & IIf(failed = "", "", ", ") & sh
                prot.Remove sh
            End If
        End If
    Next sh
    If failed <> "" Then
        ReprotectAll wb, prot
        Finish report, "STOPPED: couldn't unprotect " & failed & "." & vbLf & _
                       "Put the sheet password in SHEET_PASSWORD at the top and run again. Nothing was renamed."
        Exit Sub
    End If

    On Error GoTo RenameFailed
    ' Step 1: temporary names, so codes swapped between sheets can't collide.
    For i = 1 To plan.Count
        p = plan(i)
        tempName = "_RENAME" & i & "_" & Mid$(p(2), 2)
        Do While taken.Exists(tempName)
            tempName = tempName & "X"
        Loop
        wb.Worksheets(p(0)).ListObjects(p(1)).Name = tempName
        temps.Add tempName
    Next i
    ' Step 2: final names.
    For i = 1 To plan.Count
        p = plan(i)
        wb.Worksheets(p(0)).ListObjects(temps(i)).Name = p(2)
    Next i
    On Error GoTo 0

    ReprotectAll wb, prot
    Finish report, "Done: " & plan.Count & " tables renamed, " & correct & " already matched. " & _
                   prot.Count & " sheet(s) were unprotected and protected again."
    Exit Sub

RenameFailed:
    errText = Err.Description
    On Error GoTo 0
    ReprotectAll wb, prot
    Finish report, "ERROR while renaming: " & errText & vbLf & _
                   "Sheets have been protected again. Run the dry run again: it will pick up any table left half-renamed."
End Sub

Private Function IsModuleSheet(ByVal s As String) As Boolean
    IsModuleSheet = (Len(s) = 7) And (UCase$(s) Like "#[0-9A-Z][0-9A-Z][0-9A-Z]SPS")
End Function

Private Function IsProtected(ByVal ws As Worksheet) As Boolean
    IsProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
End Function

Private Function SaveProtection(ByVal ws As Worksheet) As Variant
    With ws.Protection
        SaveProtection = Array(ws.ProtectDrawingObjects, ws.ProtectContents, ws.ProtectScenarios, _
            .AllowFormattingCells, .AllowFormattingColumns, .AllowFormattingRows, _
            .AllowInsertingColumns, .AllowInsertingRows, .AllowInsertingHyperlinks, _
            .AllowDeletingColumns, .AllowDeletingRows, .AllowSorting, .AllowFiltering, _
            .AllowUsingPivotTables, ws.EnableSelection)
    End With
End Function

Private Sub ReprotectAll(ByVal wb As Workbook, ByVal prot As Object)
    Dim sh As Variant, s As Variant, ws As Worksheet
    For Each sh In prot.Keys
        Set ws = wb.Worksheets(sh)
        s = prot(sh)
        ws.Protect Password:=SHEET_PASSWORD, DrawingObjects:=s(0), Contents:=s(1), Scenarios:=s(2), _
            AllowFormattingCells:=s(3), AllowFormattingColumns:=s(4), AllowFormattingRows:=s(5), _
            AllowInsertingColumns:=s(6), AllowInsertingRows:=s(7), AllowInsertingHyperlinks:=s(8), _
            AllowDeletingColumns:=s(9), AllowDeletingRows:=s(10), AllowSorting:=s(11), _
            AllowFiltering:=s(12), AllowUsingPivotTables:=s(13)
        ws.EnableSelection = s(14)
    Next sh
End Sub

' Report (but don't change) a Module Code in the Meta table that differs from the sheet name.
Private Sub CheckMetaCode(ByVal lo As ListObject, ByVal sheetName As String, ByVal notes As Collection)
    Dim r As Long, code As String
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For r = 1 To lo.DataBodyRange.Rows.Count
        If LCase$(Trim$(lo.DataBodyRange.Cells(r, 1).Text)) = "module code" Then
            code = Trim$(lo.DataBodyRange.Cells(r, 2).Text)
            If code <> "" And UCase$(code) <> UCase$(sheetName) Then
                notes.Add sheetName & ": the Meta table still says Module Code """ & code & """. Update it by hand."
            End If
            Exit Sub
        End If
    Next r
End Sub

Private Sub Finish(ByVal report As String, ByVal summary As String)
    Dim shown As String
    Debug.Print summary
    shown = report
    If Len(shown) > 700 Then shown = Left$(shown, 700) & "..." & vbLf & "(full list in the Immediate window: Ctrl+G in the VBA editor)"
    MsgBox summary & vbLf & vbLf & shown, vbInformation, "Fix table names"
End Sub
