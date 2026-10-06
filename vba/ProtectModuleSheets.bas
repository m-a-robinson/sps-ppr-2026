Attribute VB_Name = "modProtectModuleSheets"
Option Explicit

' Protects every module sheet (named like 4001SPS) in the ACTIVE workbook so that:
'   - every cell inside a table can be edited, and the dropdowns keep working;
'   - everything outside the tables is locked.
' Table header rows and formula cells (the Hours total) stay locked, because the queries read
' the column headings by name and the total is calculated. Set LOCK_HEADERS or LOCK_FORMULAS
' to False to make those editable too.
'
' Query output sheets (FS_Meta, SES_Mapping ...) are never protected: protection blocks Refresh.
' LISTS, the source of the dropdowns, is locked completely unless PROTECT_LISTS = False.
' Sheets that are already protected are unprotected with SHEET_PASSWORD and protected again
' with these settings. UnprotectModuleSheets undoes it.
'
' Note: tables can't gain rows on a protected sheet. Unprotect first if a module needs more.

Private Const SHEET_PASSWORD As String = ""    ' leave blank for no password
Private Const PROTECT_LISTS As Boolean = True  ' also lock LISTS completely
Private Const LOCK_HEADERS As Boolean = True   ' keep table header rows locked
Private Const LOCK_FORMULAS As Boolean = True  ' keep formula cells (Hours total) locked

Public Sub ProtectModuleSheets()
    Dim ws As Worksheet, lo As ListObject, c As Range
    Dim nDone As Long, nCells As Long, failed As String

    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
        If InScope(ws) Then
            If IsProtected(ws) Then
                On Error Resume Next
                ws.Unprotect SHEET_PASSWORD
                On Error GoTo 0
            End If
            If IsProtected(ws) Then
                failed = failed & IIf(failed = "", "", ", ") & ws.Name
            Else
                ' Lock everything, then open up the table cells.
                ws.Cells.Locked = True
                If IsModuleSheet(ws.Name) Then
                    For Each lo In ws.ListObjects
                        If Not lo.DataBodyRange Is Nothing Then
                            lo.DataBodyRange.Locked = False
                            nCells = nCells + lo.DataBodyRange.Cells.Count
                            If LOCK_FORMULAS Then
                                For Each c In lo.DataBodyRange.Cells
                                    If c.HasFormula Then
                                        c.Locked = True
                                        nCells = nCells - 1
                                    End If
                                Next c
                            End If
                        End If
                        If Not LOCK_HEADERS And lo.ShowHeaders Then lo.HeaderRowRange.Locked = False
                    Next lo
                End If

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
    Application.ScreenUpdating = True

    MsgBox "Protected " & nDone & " sheet(s): " & nCells & " table cells are editable, everything else is locked." & _
           IIf(failed = "", "", vbLf & vbLf & "Couldn't unprotect " & failed & " to update it: put the password in SHEET_PASSWORD."), _
           vbInformation, "Protect module sheets"
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
