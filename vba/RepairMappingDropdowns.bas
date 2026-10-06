Attribute VB_Name = "modRepairMappingDropdowns"
Option Explicit

' Repairs the Graduate Skills dropdowns in every module _Mapping table of the
' ACTIVE workbook, in one pass.
'
'   DRY_RUN = True   changes nothing; opens a new workbook listing every planned change
'   DRY_RUN = False  makes the changes, then opens the same list marked Done
'
' Fixes automatically:
'   1. The four misspelt skills in the LISTS skills table (and any mapping row still using them).
'   2. Sheet-level copies of the dropdown names (SkillNameMap, ScientificKnowledge ...) left
'      behind when a sheet was copied from another workbook. They point at #REF! or another
'      file and override the good names, so the Skill dropdown on that sheet stops working.
'   3. Links to other workbooks created by those copied names.
'   4. Missing or broken dropdowns: any mapping table with a gap gets the Category list
'      (=LISTS!$J$7:$J$16) and the dependent Skill list on every row.
'   5. Category text that doesn't exactly match the list ("&" for "and", capitals).
'   6. Skills filed under the wrong category: the category is changed to the one the skill
'      belongs to in LISTS (every skill belongs to exactly one category).
'   7. Assessment numbers written AS1, AS2 -> A1, A2.
'   8. Skill wording that differs only in capitals, hyphens or "&" from the list.
' Reports only (needs a person): skills not in the list, and Assessment No / LO Number cells
' that aren't in A1 / LO1 form (notes, "A1, A2", pasted PLO blocks).
'
' Protected sheets are unprotected only for the change and protected again with their
' original settings. Changing a Category never clears the Skill in Excel: after this runs,
' use Data > Data Validation > Circle Invalid Data to spot any row still out of step.

Private Const DRY_RUN As Boolean = True
Private Const SHEET_PASSWORD As String = ""   ' sheet protection password, if any
Private Const CATEGORY_LIST As String = "=LISTS!$J$7:$J$16"
Private Const DROPDOWN_NAMES As String = "|SkillNameMap|ScientificKnowledge|TechnicalSkills|ApplicationofKnowledge|UnderstandingResearch|SelfEvaluation|Communication|ProblemSolving|MgmtSelf|DeliveryEnvironment|ProfRelationships|"

' Each plan item: Array(kind, sheet, item, fromText, toText, what, object)
' kind = VALUE | NAME | LINK | DV | REPORT
Private plan As Collection
Private spell As Object       ' old spelling -> correct spelling
Private catCanon As Object    ' normalised category -> text in LISTS!J7:J16
Private skillCanon As Object  ' normalised skill -> text in the skills table
Private skillHome As Object   ' normalised skill -> its category (list text)

Public Sub RepairMappingDropdowns()
    Dim wb As Workbook, ws As Worksheet, listsWs As Worksheet, lo As ListObject, lt As ListObject
    Dim nm As Name, startSheet As Object, prot As Object, toChange As Object
    Dim p As Variant, sh As Variant, links As Variant, stat() As String
    Dim i As Long, r As Long, colCat As Long, colSkill As Long
    Dim c As String, s As String, shortName As String, failed As String, errText As String
    Dim foundNames As String

    Set wb = ActiveWorkbook
    Set plan = New Collection
    Set spell = CreateObject("Scripting.Dictionary"): spell.CompareMode = vbTextCompare
    spell("Domian-specfic analytical skills") = "Domain-specific analytical skills"
    spell("Reserch literacy") = "Research literacy"
    spell("Evidne-based practice") = "Evidence-based practice"
    spell("Interpesronal communication") = "Interpersonal communication"
    Set catCanon = CreateObject("Scripting.Dictionary")
    Set skillCanon = CreateObject("Scripting.Dictionary")
    Set skillHome = CreateObject("Scripting.Dictionary")

    ' --- LISTS: categories and the skills table -------------------------------------
    On Error Resume Next
    Set listsWs = wb.Worksheets("LISTS")
    If Not listsWs Is Nothing Then Set lt = listsWs.ListObjects("LISTS_Table")
    On Error GoTo 0
    If lt Is Nothing Then
        MsgBox "No LISTS sheet with a LISTS_Table in " & wb.Name & ". Nothing was checked.", vbExclamation, "Repair mapping dropdowns"
        Exit Sub
    End If
    For r = 7 To 16
        c = CellText(listsWs.Cells(r, "J"))
        If c <> "" Then catCanon(Norm(c)) = c
    Next r
    colCat = ColIndex(lt, "Category")
    colSkill = ColIndex(lt, "Skill")
    If colCat = 0 Or colSkill = 0 Or lt.DataBodyRange Is Nothing Then
        MsgBox "LISTS_Table needs Category and Skill columns. Nothing was checked.", vbExclamation, "Repair mapping dropdowns"
        Exit Sub
    End If
    For r = 1 To lt.DataBodyRange.Rows.Count
        c = CellText(lt.DataBodyRange.Cells(r, colCat))
        s = CellText(lt.DataBodyRange.Cells(r, colSkill))
        If c <> "" And s <> "" Then
            If spell.Exists(s) Then
                AddItem "VALUE", "LISTS", lt.DataBodyRange.Cells(r, colSkill).Address(False, False), s, spell(s), "Correct spelling in the skills list"
                s = spell(s)
            End If
            If catCanon.Exists(Norm(c)) Then
                skillCanon(Norm(s)) = s
                skillHome(Norm(s)) = catCanon(Norm(c))
            Else
                AddItem "REPORT", "LISTS", lt.DataBodyRange.Cells(r, colCat).Address(False, False), c, "", "Category in the skills table isn't in LISTS!J7:J16"
            End If
        End If
    Next r

    ' --- Names: sheet-level copies to delete, workbook-level ones to check ------------
    For Each nm In wb.Names
        shortName = Mid$(nm.Name, InStrRev(nm.Name, "!") + 1)
        If InStr(1, DROPDOWN_NAMES, "|" & shortName & "|", vbTextCompare) > 0 Then
            If TypeName(nm.Parent) = "Worksheet" Then
                AddItem "NAME", nm.Parent.Name, shortName, nm.RefersTo, "", "Delete sheet-level copy of a dropdown name", nm
            Else
                foundNames = foundNames & "|" & shortName & "|"
                If InStr(nm.RefersTo, "#REF") > 0 Or InStr(nm.RefersTo, "[") > 0 Then
                    AddItem "REPORT", "", shortName, nm.RefersTo, "", "Workbook-level dropdown name is broken: fix it in Name Manager"
                End If
            End If
        End If
    Next nm
    For Each sh In Split(Mid$(DROPDOWN_NAMES, 2, Len(DROPDOWN_NAMES) - 2), "|")
        If InStr(1, foundNames, "|" & sh & "|", vbTextCompare) = 0 Then
            AddItem "REPORT", "", CStr(sh), "", "", "Workbook-level dropdown name is missing: add it in Name Manager"
        End If
    Next sh

    ' --- Links to other workbooks ------------------------------------------------------
    links = wb.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        For i = LBound(links) To UBound(links)
            AddItem "LINK", "", CStr(links(i)), "", "", "Break link to another workbook"
        Next i
    End If

    ' --- Every module _Mapping table --------------------------------------------------
    For Each ws In wb.Worksheets
        If IsModuleSheet(ws.Name) Then
            For Each lo In ws.ListObjects
                If LCase$(Right$(lo.DisplayName, 8)) = "_mapping" Then CheckMappingTable ws, lo
            Next lo
        End If
    Next ws

    If plan.Count = 0 Then
        MsgBox "Nothing to repair in " & wb.Name & ".", vbInformation, "Repair mapping dropdowns"
        Exit Sub
    End If
    If DRY_RUN Then
        WriteReport wb, Empty
        Exit Sub
    End If

    ' --- Apply ------------------------------------------------------------------------
    Set startSheet = wb.ActiveSheet
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Unprotect every sheet that will change, remembering its settings.
    Set prot = CreateObject("Scripting.Dictionary")
    Set toChange = CreateObject("Scripting.Dictionary")
    For i = 1 To plan.Count
        p = plan(i)
        If p(0) = "VALUE" Or p(0) = "DV" Then toChange(p(1)) = True
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
        RestoreApp
        MsgBox "STOPPED: couldn't unprotect " & failed & "." & vbLf & _
               "Put the sheet password in SHEET_PASSWORD at the top and run again. Nothing was changed.", vbExclamation, "Repair mapping dropdowns"
        Exit Sub
    End If

    ReDim stat(1 To plan.Count)
    On Error GoTo ApplyFailed
    ' Pass 1: cell values (LISTS spellings first, so the skill lists recalculate).
    For i = 1 To plan.Count
        p = plan(i)
        If p(0) = "VALUE" Then
            wb.Worksheets(p(1)).Range(p(2)).Value = p(4)
            stat(i) = "Done"
        ElseIf p(0) = "REPORT" Then
            stat(i) = "Fix by hand"
        End If
    Next i
    ' Pass 2: delete the sheet-level name copies.
    For i = 1 To plan.Count
        p = plan(i)
        If p(0) = "NAME" Then
            p(6).Delete
            stat(i) = "Done"
        End If
    Next i
    ' Pass 3: dropdowns.
    For i = 1 To plan.Count
        p = plan(i)
        If p(0) = "DV" Then stat(i) = ApplyDropdowns(wb.Worksheets(p(1)), CStr(p(2)), CStr(p(3)))
    Next i
    ' Pass 4: break links that are no longer needed (Excel may already have dropped them).
    For i = 1 To plan.Count
        p = plan(i)
        If p(0) = "LINK" Then
            On Error Resume Next
            wb.BreakLink Name:=p(2), Type:=xlLinkTypeExcelLinks
            stat(i) = IIf(Err.Number = 0, "Done", "Already gone")
            Err.Clear
            On Error GoTo ApplyFailed
        End If
    Next i
    On Error GoTo 0

    ReprotectAll wb, prot
    startSheet.Activate
    RestoreApp
    WriteReport wb, stat
    Exit Sub

ApplyFailed:
    errText = Err.Description
    On Error GoTo 0
    ReprotectAll wb, prot
    startSheet.Activate
    RestoreApp
    For i = 1 To plan.Count
        If stat(i) = "" Then stat(i) = "Not done (stopped)"
    Next i
    WriteReport wb, stat
    MsgBox "STOPPED part-way: " & errText & vbLf & "Sheets were protected again. The report shows what was done; run the dry run again to see what is left.", vbExclamation, "Repair mapping dropdowns"
End Sub

' Plan every change for one mapping table.
Private Sub CheckMappingTable(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim cA As Long, cL As Long, cC As Long, cS As Long, r As Long
    Dim a As String, lon As String, cat As String, sk As String, newCat As String, ns As String
    Dim why As String, needDV As Boolean, catCell As Range, skCell As Range, cell As Range

    If lo.DataBodyRange Is Nothing Then Exit Sub
    cA = ColIndex(lo, "Assessment No")
    cL = ColIndex(lo, "LO Number")
    cC = ColIndex(lo, "Graduate Skills Category")
    cS = ColIndex(lo, "Graduate Skill")
    If cC = 0 Or cS = 0 Then
        AddItem "REPORT", ws.Name, lo.Range.Address(False, False), "", "", "Mapping table has no Graduate Skills Category / Graduate Skill columns"
        Exit Sub
    End If

    For r = 1 To lo.DataBodyRange.Rows.Count
        Set catCell = lo.DataBodyRange.Cells(r, cC)
        Set skCell = lo.DataBodyRange.Cells(r, cS)
        If Not HasListRule(catCell, "J$7") Or Not HasListRule(skCell, "SkillNameMap") Then needDV = True

        If cA > 0 Then
            Set cell = lo.DataBodyRange.Cells(r, cA)
            a = CellText(cell)
            If UCase$(a) Like "AS#" Or UCase$(a) Like "AS##" Then
                AddItem "VALUE", ws.Name, cell.Address(False, False), a, "A" & Mid$(a, 3), "Assessment number in A1 form"
            ElseIf a <> "" And Not (UCase$(a) Like "A#" Or UCase$(a) Like "A##") Then
                AddItem "REPORT", ws.Name, cell.Address(False, False), a, "", "Assessment No isn't a single A1/A2: fix by hand (one item per row)"
            End If
        End If
        If cL > 0 Then
            Set cell = lo.DataBodyRange.Cells(r, cL)
            lon = CellText(cell)
            If lon <> "" And Not (UCase$(lon) Like "LO#" Or UCase$(lon) Like "LO##") Then
                AddItem "REPORT", ws.Name, cell.Address(False, False), lon, "", "LO Number isn't a single LO1/LO2: fix by hand (one LO per row)"
            End If
        End If

        cat = CellText(catCell)
        sk = CellText(skCell)

        ' Skill: spelling, then wording that differs only in capitals, hyphens or "&".
        If sk <> "" Then
            If spell.Exists(sk) Then
                AddItem "VALUE", ws.Name, skCell.Address(False, False), sk, spell(sk), "Correct skill spelling"
                sk = spell(sk)
            ElseIf skillCanon.Exists(Norm(sk)) Then
                If skillCanon(Norm(sk)) <> sk Then
                    AddItem "VALUE", ws.Name, skCell.Address(False, False), sk, skillCanon(Norm(sk)), "Skill wording matched to the list"
                    sk = skillCanon(Norm(sk))
                End If
            End If
        End If

        ' Category: exact list wording, then the category the skill belongs to.
        newCat = cat
        why = ""
        If cat <> "" Then
            If catCanon.Exists(Norm(cat)) Then
                If catCanon(Norm(cat)) <> cat Then
                    newCat = catCanon(Norm(cat))
                    why = "Category wording matched to the list"
                End If
            End If
        End If
        ns = Norm(sk)
        If sk <> "" Then
            If skillHome.Exists(ns) Then
                If Norm(newCat) <> Norm(skillHome(ns)) Then
                    newCat = skillHome(ns)
                    why = IIf(cat = "", "Category added to match the skill", "Category changed to the one '" & sk & "' belongs to")
                End If
            Else
                AddItem "REPORT", ws.Name, skCell.Address(False, False), sk, "", "Skill isn't in the list: pick one from the dropdown"
            End If
        End If
        If newCat <> cat Then
            AddItem "VALUE", ws.Name, catCell.Address(False, False), cat, newCat, why
        ElseIf cat <> "" And Not catCanon.Exists(Norm(cat)) Then
            AddItem "REPORT", ws.Name, catCell.Address(False, False), cat, "", "Category isn't in the list: pick one from the dropdown"
        End If
    Next r

    If needDV Then
        AddItem "DV", ws.Name, lo.ListColumns(cC).DataBodyRange.Address(False, False), _
                lo.ListColumns(cS).DataBodyRange.Address(False, False), "", "Re-apply the Category and Skill dropdowns"
    End If
End Sub

' Category list on the category column; dependent Skill list on the skill column.
Private Function ApplyDropdowns(ByVal ws As Worksheet, ByVal catAddr As String, ByVal skAddr As String) As String
    Dim catRng As Range, skRng As Range, c As Range, firstCat As String

    Set catRng = ws.Range(catAddr)
    Set skRng = ws.Range(skAddr)
    With catRng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=CATEGORY_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' A relative reference in a validation formula is read from the active cell,
    ' so select the first skill cell and write the rule for that row ($E46 style).
    firstCat = catRng.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=True)
    On Error Resume Next
    Application.Goto skRng.Cells(1, 1)
    If Err.Number = 0 Then
        On Error GoTo 0
        With skRng.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                 Formula1:="=INDIRECT(VLOOKUP(" & firstCat & ",SkillNameMap,2,FALSE))"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        ApplyDropdowns = "Done"
    Else
        ' Sheet can't be selected (e.g. hidden): one rule per cell with a fixed reference.
        On Error GoTo 0
        For Each c In skRng.Cells
            With c.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                     Formula1:="=INDIRECT(VLOOKUP(" & ws.Cells(c.Row, catRng.Column).Address & ",SkillNameMap,2,FALSE))"
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        Next c
        ApplyDropdowns = "Done (per-cell rules)"
    End If
End Function

Private Sub AddItem(ByVal kind As String, ByVal sheetName As String, ByVal item As String, _
                    ByVal fromText As String, ByVal toText As String, ByVal what As String, Optional ByVal obj As Object)
    plan.Add Array(kind, sheetName, item, fromText, toText, what, obj)
End Sub

' True if the cell has a list rule whose formula contains the token and no broken reference.
Private Function HasListRule(ByVal c As Range, ByVal token As String) As Boolean
    Dim f As String
    On Error Resume Next
    f = c.Validation.Formula1
    On Error GoTo 0
    HasListRule = (InStr(1, f, token, vbTextCompare) > 0) And (InStr(f, "[") = 0) And (InStr(f, "#REF") = 0)
End Function

Private Function ColIndex(ByVal lo As ListObject, ByVal header As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If LCase$(Trim$(lc.Name)) = LCase$(header) Then
            ColIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function CellText(ByVal c As Range) As String
    If IsError(c.Value) Then Exit Function
    CellText = Trim$(Replace(CStr(c.Value), Chr$(160), " "))
End Function

' Lower case, "&" read as "and", hyphens as spaces, single spaces.
Private Function Norm(ByVal s As String) As String
    s = LCase$(Replace(s, Chr$(160), " "))
    s = Replace(s, "&", " and ")
    s = Replace(s, "-", " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = Trim$(s)
End Function

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
    If prot Is Nothing Then Exit Sub
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

Private Sub RestoreApp()
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Writes the plan (dry run) or the results into a new workbook.
Private Sub WriteReport(ByVal wb As Workbook, ByVal stat As Variant)
    Dim rep As Workbook, sh As Worksheet, i As Long, p As Variant, row As Long
    Dim nChange As Long, nManual As Long, status As String, item As String

    Set rep = Workbooks.Add(xlWBATWorksheet)
    Set sh = rep.Worksheets(1)
    sh.Name = "Repair report"
    sh.Range("A1").Value = IIf(IsEmpty(stat), "DRY RUN: nothing has been changed in ", "Changes made to ") & wb.Name
    sh.Range("A3:F3").Value = Array("Sheet", "Cell / item", "What", "Now", "Becomes", "Status")
    row = 3
    For i = 1 To plan.Count
        p = plan(i)
        row = row + 1
        item = p(2)
        If p(0) = "DV" Then item = p(2) & " and " & p(3)
        If p(0) = "REPORT" Then
            status = "Fix by hand"
            nManual = nManual + 1
        ElseIf IsEmpty(stat) Then
            status = "Planned"
            nChange = nChange + 1
        Else
            status = stat(i)
            nChange = nChange + 1
        End If
        sh.Cells(row, 1).Value = p(1)
        sh.Cells(row, 2).Value = item
        sh.Cells(row, 3).Value = p(5)
        sh.Cells(row, 4).Value = "'" & IIf(p(0) = "DV", "", p(3))
        sh.Cells(row, 5).Value = "'" & p(4)
        sh.Cells(row, 6).Value = status
    Next i
    sh.Range("A2").Value = nChange & " automatic fix(es), " & nManual & " item(s) to fix by hand." & _
        IIf(IsEmpty(stat), " Set DRY_RUN = False and run again on " & wb.Name & " to apply.", "")
    sh.Range("A3:F3").Font.Bold = True
    sh.Range("A3:F" & row).AutoFilter
    sh.Columns("A:B").ColumnWidth = 14
    sh.Columns("C").ColumnWidth = 55
    sh.Columns("D:E").ColumnWidth = 40
    sh.Columns("F").ColumnWidth = 18
    sh.Range("A3:F" & row).WrapText = True
    sh.Range("A3:F" & row).VerticalAlignment = xlTop
End Sub
