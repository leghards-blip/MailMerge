Attribute VB_Name = "DriveDropdown_F5_v1"

Option Explicit

' ===== DriveDropdown_G5_v1.bas =====
' Purpose:
'   1) Detect available drive roots (C:\ .. Z:\) on the local machine.
'   2) Create/refresh a very hidden config sheet to store the list.
'   3) Apply Data Validation (list) to a target cell (F5) with those roots.
'   4) Provide helpers to read the chosen value and ensure it has a trailing "\".
'
' Public macros:
'   SetupDriveDropdown_G5()                 ' Creates the dropdown in ActiveSheet!G5
'   RefreshDriveDropdown_G5()               ' Refreshes the list & validation (re-scan drives)
'
' Public helpers:
'   GetRootFromF5(Optional ws As Worksheet) As String
'   NormalizeRoot(ByVal root As String) As String
'
' Notes:
'   - The drive list is stored on a very hidden sheet: "__Config".
'   - Named range created/updated: DriveList
'   - Preferred default selection: Z:\ then O:\ then first detected drive.
'   - If you prefer to pin this to a specific sheet, change the code to reference it by name.
'
' Integration with your path formula:
'   Change your path-building formula to use $G$5 as the root, e.g.:
'     =LET(root; $G$5; ... root & "Devon\" & ... )
'
Private Const CFG_SHEET As String = "__Config"
Private Const NAME_DRIVELIST As String = "DriveList"

Public Sub SetupDriveDropdown_G5()
    ' Create/refresh the list and apply data validation to F5 on the active sheet.
    Dim wsTarget As Worksheet
    Set wsTarget = ActiveSheet
    ApplyDriveDropdown wsTarget.Range("G5")
End Sub

Public Sub RefreshDriveDropdown_G5()
    ' Refresh list & validation, keeping the current sheet's F5 as the target.
    SetupDriveDropdown_G5
End Sub

Private Sub ApplyDriveDropdown(ByVal target As Range)
    Dim wsCfg As Worksheet
    Dim roots As Collection, d As Variant
    Dim last As Long, pref As String
    
    Set wsCfg = EnsureConfigSheet()
    Set roots = DetectDrives()
    
    ' Ensure Z:\ and O:\ are included at top preference if present
    ' We will write the detected list, but pick default by preference.
    wsCfg.Cells.ClearContents
    wsCfg.Range("A1").Value = "Drive roots"
    
    last = 1
    For Each d In roots
        last = last + 1
        wsCfg.Cells(last, 1).Value = d
    Next d
    
    ' Define/refresh a named range for validation
    On Error Resume Next
    ThisWorkbook.Names(NAME_DRIVELIST).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=NAME_DRIVELIST, _
        RefersTo:=wsCfg.Range(wsCfg.Cells(2, 1), wsCfg.Cells(last, 1))
    
    ' Apply Data Validation to G5
    With target
        .ClearContents
        On Error Resume Next
        .Validation.Delete
        On Error GoTo 0
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=" & NAME_DRIVELIST
        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
        .Validation.InputTitle = "Choose data drive"
        .Validation.InputMessage = "Pick the mapped/local drive root (e.g., Z:\, O:\)."
        .Validation.ErrorTitle = "Invalid selection"
        .Validation.ErrorMessage = "Please choose a drive from the list."
        .Validation.ShowError = True
        .Validation.ShowInput = True
    End With
    
    ' Choose a sensible default: Z:\ if present, otherwise O:\, otherwise first list item
    pref = PickPreferredDefault(wsCfg.Range("A2:A" & last))
    If Len(pref) > 0 Then
        target.Value = pref
    End If
    
    ' Keep config sheet very hidden
    wsCfg.Visible = xlSheetVeryHidden
End Sub

Private Function EnsureConfigSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CFG_SHEET
    End If
    ' Make sure it's very hidden (not visible in Unhide dialog)
    ws.Visible = xlSheetVeryHidden
    Set EnsureConfigSheet = ws
End Function

Private Function DetectDrives() As Collection
    Dim col As New Collection
    Dim d As Integer
    Dim root As String
    
    On Error Resume Next
    For d = 67 To 90 ' C through Z
        root = Chr$(d) & ":\\"
        root = Left$(root, 3) ' "C:\"
        If Len(Dir(root, vbDirectory)) > 0 Then
            col.Add root
        End If
    Next d
    On Error GoTo 0
    
    Set DetectDrives = col
End Function

Private Function PickPreferredDefault(ByVal rng As Range) As String
    Dim c As Range
    ' Prefer Z:\ then O:\, else first non-blank in the list
    For Each c In rng.Cells
        If UCase$(c.Value) = "Z:\" Then
            PickPreferredDefault = "Z:\"
            Exit Function
        End If
    Next c
    For Each c In rng.Cells
        If UCase$(c.Value) = "O:\" Then
            PickPreferredDefault = "O:\"
            Exit Function
        End If
    Next c
    For Each c In rng.Cells
        If Len(c.Value) > 0 Then
            PickPreferredDefault = CStr(c.Value)
            Exit Function
        End If
    Next c
End Function

Public Function GetRootFromG5(Optional ws As Worksheet) As String
    Dim v As String
    If ws Is Nothing Then Set ws = ActiveSheet
    v = CStr(ws.Range("G5").Value)
    GetRootFromG5 = NormalizeRoot(v)
End Function

Public Function NormalizeRoot(ByVal root As String) As String
    root = Trim$(root)
    If Len(root) = 0 Then Exit Function
    If Len(root) = 1 Then root = UCase$(root) & ":"
    If Len(root) = 2 And Mid$(root, 2, 1) = ":" Then root = UCase$(Left$(root, 1)) & ":"
    If Right$(root, 1) <> "\" Then root = root & "\"
    NormalizeRoot = root
End Function
