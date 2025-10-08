Attribute VB_Name = "GmailSMTP_Levy_v1"
Option Explicit

' ===== GmailSMTP_Levy_v1.bas (Updated by Gemini) =====
' Send emails as levy@beraderproperties.com via Gmail SMTP
' Now includes dual-port fallback (587/TLS, 465/SSL) and improved error handling.
' Works independently of Outlook profiles or who runs the macro.
'
' Public:
'   SetLevyCredentials()     - Prompt & store levy@ username + App Password (very-hidden, obfuscated)
'   ClearLevyCredentials()   - Remove stored creds
'   SendLevyEmail_CDO(...)   - Send message with optional CC/BCC and attachments
'
' Prereq:
'   - On levy@ Google account: enable 2-Step Verification, create "App Password" for Mail.
'   - Windows with CDO (present by default on most systems). Late-bound; no references needed.
'
Private Const CFG_SHEET As String = "__Config"
Private Const NM_USER As String = "LEVY_SMTP_USER"
Private Const NM_PASS As String = "LEVY_SMTP_APP"
Private Const SMTP_HOST As String = "smtp-relay.gmail.com"

' === Public API ===============================================================

Public Sub SetLevyCredentials()
    Dim u As String, p As String
    u = InputBox("Enter Gmail username:", "Levy SMTP setup", "levy@beraderproperties.com")
    If Len(u) = 0 Then Exit Sub
    p = InputBox("Enter App Password (16 chars, no spaces):", "Levy SMTP setup")
    If Len(p) = 0 Then Exit Sub
    SaveObfuscated NM_USER, u
    SaveObfuscated NM_PASS, p
    MsgBox "Credentials saved for levy@ SMTP (stored obfuscated on __Config).", vbInformation
End Sub

Public Sub ClearLevyCredentials()
    On Error Resume Next
    ThisWorkbook.Names(NM_USER).Delete
    ThisWorkbook.Names(NM_PASS).Delete
    EnsureConfigSheet.Visible = xlSheetVeryHidden
    On Error GoTo 0
    MsgBox "Levy SMTP credentials cleared.", vbInformation
End Sub

' Send as levy@ via Gmail SMTP using CDO.
' Returns True if sent, False otherwise.
Public Function SendLevyEmail_CDO(ByVal toList As String, ByVal subject As String, _
                                  Optional ByVal htmlBody As String = "", _
                                  Optional ByVal textBody As String = "", _
                                  Optional ByVal ccList As String = "", _
                                  Optional ByVal bccList As String = "", _
                                  Optional attachments As Variant, _
                                  Optional ByVal replyTo As String = "") As Boolean
    Dim user As String, pwd As String
    Dim cfg As Object, msg As Object
    Dim i As Long
    Dim portsToTry As Variant, p As Variant
    Dim lastError As String, success As Boolean

    user = LoadObfuscated(NM_USER)
    ' pwd = LoadObfuscated(NM_PASS) ' No longer needed for authentication
    
    ' We still need the username for the 'From' address
    If Len(user) = 0 Then
        MsgBox "Levy SMTP username not set. Run SetLevyCredentials first (password will be ignored but username is needed).", vbExclamation
        Exit Function
    End If

    portsToTry = Array(587, 465)
    success = False

    For Each p In portsToTry
        On Error GoTo SendAttemptFailed
        Set cfg = CreateObject("CDO.Configuration")
        With cfg.Fields
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_HOST
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(p)
            
            ' --- MODIFIED: Use Anonymous authentication for IP-based relay ---
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0 ' 0 = cdoAnonymous
            
            ' These fields are no longer sent
            ' .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = user
            ' .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = pwd
            
            ' Encryption is still required by the server rule
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
            
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
            .Update
        End With

        Set msg = CreateObject("CDO.Message")
        Set msg.Configuration = cfg
        With msg
            .From = """" & "Levy Statements" & """ <" & user & ">"
            .To = toList
            If Len(ccList) > 0 Then .CC = ccList
            If Len(bccList) > 0 Then .BCC = bccList
            If Len(replyTo) > 0 Then .ReplyTo = replyTo
            .Subject = subject
            If Len(htmlBody) > 0 Then
                .HTMLBody = htmlBody
            Else
                .TextBody = textBody
            End If
            If Not Ismissing(attachments) Then
                If IsArray(attachments) Then
                    For i = LBound(attachments) To UBound(attachments)
                        If Len(attachments(i)) > 0 Then .AddAttachment CStr(attachments(i))
                    Next i
                ElseIf VarType(attachments) = vbString Then
                    If Len(attachments) > 0 Then .AddAttachment CStr(attachments)
                End If
            End If
            .Send
        End With
        
        success = True
        Exit For

SendAttemptFailed:
        If Err.Number <> 0 Then
            lastError = lastError & "Port " & p & " failed: " & Err.Description & vbCrLf
            Err.Clear
        End If
    Next p

    If success Then
        SendLevyEmail_CDO = True
    Else
        SendLevyEmail_CDO = False
        MsgBox "Send failure (levy@ Gmail SMTP):" & vbCrLf & vbCrLf & Trim(lastError), vbCritical
    End If

End Function


' === Credential storage (very hidden + simple obfuscation) ====================
' (No changes in this section)

Private Sub SaveObfuscated(ByVal nameKey As String, ByVal plain As String)
    Dim ws As Worksheet, obf As String
    Set ws = EnsureConfigSheet()
    obf = XorObfuscate(plain, "k3y!Levy2025")
    On Error Resume Next
    ThisWorkbook.Names(nameKey).Delete
    On Error GoTo 0
    
    ' --- FIX: Escape any double-quotes in the obfuscated string to prevent formula errors ---
    ThisWorkbook.Names.Add Name:=nameKey, RefersTo:="=""" & Replace(obf, """", """""") & """"
    
    ws.Visible = xlSheetVeryHidden
End Sub

Private Function LoadObfuscated(ByVal nameKey As String) As String
    Dim v As String
    On Error Resume Next
    v = CStr(Evaluate(nameKey))
    On Error GoTo 0
    If Len(v) = 0 Then Exit Function
    LoadObfuscated = XorObfuscate(v, "k3y!Levy2025")
End Function

Private Function XorObfuscate(ByVal s As String, ByVal k As String) As String
    Dim i As Long, kc As Long, ch As Integer, res As String
    If Len(k) = 0 Then XorObfuscate = s: Exit Function
    kc = Len(k)
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1)) Xor Asc(Mid$(k, ((i - 1) Mod kc) + 1, 1))
        res = res & Chr$(ch)
    Next i
    XorObfuscate = res
End Function

Private Function EnsureConfigSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CFG_SHEET
    End If
    ws.Visible = xlSheetVeryHidden
    Set EnsureConfigSheet = ws
End Function