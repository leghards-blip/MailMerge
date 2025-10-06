
Attribute VB_Name = "SendLevyStatements_v3c"
Option Explicit

' ===== SendLevyStatements_v3c.bas =====
' Column layout (Statements):
'   A: Email Address
'   B: Unit Number
'   C: PDF File Path (Auto)
'   D: File Exists?           (owned by FileCheck; e.g., "?? Found", "Found", TRUE, "Yes")
'   E: Email Status           (this macro writes here only)
'   F1: Complex Code
'   F2: Month + Year (e.g., "AUG 2025")
'   F6: (Optional) Template Path override; otherwise defaults to <ThisWorkbook.Path>\email_template.html
'
' Button macro:
'   SendAllLevyStatements_UsingTemplate_V3c
'
Private Const SHEET_NAME As String = "Statements"
Private Const COL_EMAIL As Long = 1   ' A
Private Const COL_UNIT  As Long = 2   ' B
Private Const COL_PATH  As Long = 3   ' C
Private Const COL_FILE  As Long = 4   ' D
Private Const COL_STAT  As Long = 5   ' E

Public Sub SendAllLevyStatements_UsingTemplate_V3c()
    Dim ws As Worksheet
    Dim r As Long, blanks As Long
    Dim emailAddr As String, unitNo As String, pdfPath As String
    Dim complexCode As String, monthYear As String, tplPath As String
    Dim subj As String, htmlBody As String
    Dim ok As Boolean, sent As Long, skipped As Long
    
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    complexCode = Trim$(CStr(ws.Range("F1").Value))
    monthYear  = Trim$(CStr(ws.Range("F2").Value))
    tplPath    = GetTemplatePath(ws)
    
    If Dir(tplPath, vbNormal) = vbNullString Then
        MsgBox "Email template not found:" & vbCrLf & tplPath & vbCrLf & _
               "Place 'email_template.html' next to this workbook or enter a full path in F6.", vbCritical
        GoTo Clean
    End If
    
    r = 2: blanks = 0: sent = 0: skipped = 0
    Do While blanks < 10
        emailAddr = Trim$(CStr(ws.Cells(r, COL_EMAIL).Value))
        unitNo    = Trim$(CStr(ws.Cells(r, COL_UNIT).Value))
        pdfPath   = Trim$(CStr(ws.Cells(r, COL_PATH).Value))
        
        If Len(emailAddr) = 0 And Len(unitNo) = 0 And Len(pdfPath) = 0 Then
            blanks = blanks + 1
        Else
            blanks = 0
            
            ' Check FileCheck status in column D
            If Not IsFileFlagOK(ws.Cells(r, COL_FILE).Value) Then
                ws.Cells(r, COL_STAT).Value = "Missing file"
                skipped = skipped + 1
                GoTo NextRow
            End If
            
            If Len(emailAddr) = 0 Then
                ws.Cells(r, COL_STAT).Value = "No email"
                skipped = skipped + 1
                GoTo NextRow
            End If
            
            subj = complexCode & " " & monthYear & " Levy Statement - " & unitNo
            htmlBody = TemplateEngine_v1.BuildEmailHtmlFromFile( _
                           tplPath, _
                           Array("UNIT","COMPLEX","MONTHYEAR"), _
                           Array(unitNo, complexCode, monthYear) _
                        )
            
            ok = GmailSMTP_Levy_v1.SendLevyEmail_CDO( _
                    toList:=emailAddr, _
                    subject:=subj, _
                    htmlBody:=htmlBody, _
                    attachments:=Array(pdfPath) _
                 )
            
            If ok Then
                ws.Cells(r, COL_STAT).Value = "Sent"
                sent = sent + 1
            Else
                If Len(ws.Cells(r, COL_STAT).Value) = 0 Then ws.Cells(r, COL_STAT).Value = "Error"
            End If
        End If
        
NextRow:
        r = r + 1
    Loop
    
    MsgBox "Levy mail run complete." & vbCrLf & "Sent: " & sent & vbCrLf & "Skipped: " & skipped, vbInformation
Clean:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    MsgBox "Error in SendAllLevyStatements_UsingTemplate_V3c: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Clean
End Sub

Private Function GetTemplatePath(ByVal ws As Worksheet) As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(ws.Range("F6").Value)) ' optional override
    On Error GoTo 0
    If Len(v) > 0 Then
        GetTemplatePath = v
    Else
        GetTemplatePath = ThisWorkbook.Path & Application.PathSeparator & "email_template.html"
    End If
End Function

Private Function IsFileFlagOK(ByVal flagValue As Variant) As Boolean
    Dim s As String
    If VarType(flagValue) = vbBoolean Then
        IsFileFlagOK = CBool(flagValue)
        Exit Function
    End If
    s = UCase$(Trim$(CStr(flagValue)))
    ' Accept strings that CONTAIN "Found" anywhere (e.g., "?? Found")
    If InStr(1, s, "FOUND", vbTextCompare) > 0 Then IsFileFlagOK = True: Exit Function
    If s = "TRUE" Or s = "YES" Or s = "OK" Or s = "✓" Or s = "CHECKED" Then
        IsFileFlagOK = True
    Else
        IsFileFlagOK = False
    End If
End Function
