Attribute VB_Name = "modProtectModuleSheets"
Option Explicit

' Protects every module sheet (named like 4001SPS) in the ACTIVE workbook using the cells'
' current Locked settings: locked cells (labels, fixed columns) can't be changed, unlocked
' cells (the input columns) stay editable, and the dropdowns keep working.
' UnprotectModuleSheets undoes it.
'
' Query output sheets (FS_Meta, SES_Mapping ...) are never protected, because a protected
' sheet blocks Refresh. Sheets that are already protected are left exactly as they are.
'
' Before protecting, it lists input cells that are locked: a locked cell in a table column
' that is otherwise mostly unlocked, usually left behind by pasting formats. Nobody can type
' in those once the sheet is protected. Set UNLOCK_GAPS = True to unlock them first.
'
' Note: tables can't gain rows on a protected sheet. Unprotect first if a module needs more.

Private Const SHEET_PASSWORD As String = ""    ' leave blank for no password
Private Const PROTECT_LISTS As Boolean = True  ' also protect LISTS, the dropdown source lists
Private Const UNLOCK_GAPS As Boolean = False   ' unlock the locked cells found in input columns

Public Sub ProtectModuleSheets()
    Dim ws As Worksheet, lo As ListObject, lc As ListColumn, c As Range
    Dim gaps As String, sheetGaps As String, summary As String, shown As String
    Dim nDone As Long, nAlready As Long, nGaps As Long, nUnlocked As Long

    For Each ws In ActiveWorkbook.Worksheets
        If InScope(ws) Then
            If IsProtected(ws) Then
                nAlready = nAlready + 1
            Else
                sheetGaps = ""
                If IsModuleSheet(ws.Name) Then
                    For Each lo In ws.ListObjects
                        If Not lo.DataBodyRange Is Nothing Then
                            For Each lc In lo.ListColumns
                                nUnlocked = 0
                                For Each c In lc.DataBodyRange.Cells
                                    If Not c.Locked Then nUnlocked = nUnlocked + 1
                                Next c
                                ' A mostly-unlocked column is an input column: report its locked cells.
                                If nUnlocked > 0 And nUnlocked * 2 >= lc.DataBodyRange.Cells.Count Then
                                    For Each c In lc.DataBodyRange.Cells
                                        If c.Locked And Not c.HasFormula Then
                                            sheetGaps = sheetGaps & " " & c.Address(False, False)
                                            nGaps = nGaps + 1
                                            If UNLOCK_GAPS Then c.Locked = False
                                        End If
                                    Next c
                                End If
                            Next lc
                        End If
                    Next lo
                End If
                If sheetGaps <> "" Then gaps = gaps & ws.Name & ":" & sheetGaps & vbLf

                ws.Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                    AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                    AllowInsertingColumns:=False, AllowInsertingRows:=False, AllowInsertingHyperlinks:=True, _
                    AllowDeletingColumns:=False, AllowDeletingRows:=False, AllowSorting:=False, _
                    AllowFiltering:=False, AllowUsingPivotTables:=False
                ws.EnableSelection = xlNoRestrictions
                nDone = nDone + 1
            End If
        End If
    Next ws

    summary = "Protected " & nDone & " sheet(s)" & _
              IIf(nAlready > 0, "; " & nAlready & " were already protected and were left as they were", "") & "."
    If nGaps > 0 Then
        summary = summary & vbLf & vbLf & nGaps & " locked cell(s) in input columns " & _
                  IIf(UNLOCK_GAPS, "were unlocked first:", "can't be edited now (set UNLOCK_GAPS = True and re-run after unprotecting to fix):")
        Debug.Print summary
        Debug.Print gaps
        shown = gaps
        If Len(shown) > 700 Then shown = Left$(shown, 700) & "..." & vbLf & "(full list in the Immediate window: Ctrl+G)"
        summary = summary & vbLf & shown
    End If
    MsgBox summary, vbInformation, "Protect module sheets"
End Sub

Public Sub UnprotectModuleSheets()
    Dim ws As Worksheet, n As Long, failed As String

    For Each ws In ActiveWorkbook.Worksheets
        If InScope(ws) And IsProtected(ws) Then
            On Error Resume Next
            ws.Unprotect SHEET_PASSWORD
            On Error GoTo 0
            If IsProtected(ws) Then
                failed = failed & IIf(failed = "", "", ", ") & ws.Name
            Else
                n = n + 1
            End If
        End If
    Next ws
    MsgBox "Unprotected " & n & " sheet(s)." & _
           IIf(failed = "", "", vbLf & "Couldn't unprotect " & failed & ": put the password in SHEET_PASSWORD."), _
           vbInformation, "Unprotect module sheets"
End Sub

Private Function InScope(ByVal ws As Worksheet) As Boolean
    InScope = IsModuleSheet(ws.Name) Or (PROTECT_LISTS And UCase$(ws.Name) = "LISTS")
End Function

Private Function IsModuleSheet(ByVal s As String) As Boolean
    IsModuleSheet = (Len(s) = 7) And (UCase$(s) Like "#[0-9A-Z][0-9A-Z][0-9A-Z]SPS")
End Function

Private Function IsProtected(ByVal ws As Worksheet) As Boolean
    IsProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
End Function
