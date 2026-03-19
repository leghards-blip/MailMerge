Attribute VB_Name = "GmailSMTP_Levy_v1"
Public Function SendLevyEmail_Outlook(ByVal toList As String, _
                                      ByVal subject As String, _
                                      Optional ByVal htmlBody As String = "", _
                                      Optional ByVal textBody As String = "", _
                                      Optional ByVal ccList As String = "", _
                                      Optional ByVal bccList As String = "", _
                                      Optional attachments As Variant, _
                                      Optional ByVal replyTo As String = "") As Boolean
    On Error GoTo Failed

    Dim olApp As Object
    Dim olMail As Object
    Dim acc As Object
    Dim i As Long

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    ' Ensure correct sending account
    For Each acc In olApp.Session.Accounts
        If LCase(acc.SmtpAddress) = "levy@beraderproperties.com" Then
            Set olMail.SendUsingAccount = acc
            Exit For
        End If
    Next acc

    With olMail
        .To = toList
        ' CC and BCC are intentionally disabled for this mail flow.
        .subject = subject

        If Len(htmlBody) > 0 Then
            .htmlBody = htmlBody
        Else
            .Body = textBody
        End If

        If Len(replyTo) > 0 Then
            .SentOnBehalfOfName = replyTo
        End If

        If Not IsMissing(attachments) Then
            If IsArray(attachments) Then
                For i = LBound(attachments) To UBound(attachments)
                    If Len(attachments(i)) > 0 Then .attachments.Add CStr(attachments(i))
                Next i
            ElseIf VarType(attachments) = vbString Then
                If Len(attachments) > 0 Then .attachments.Add CStr(attachments)
            End If
        End If

        .Send
    End With

    SendLevyEmail_Outlook = True
    Exit Function

Failed:
    SendLevyEmail_Outlook = False
    MsgBox "Outlook send failed:" & vbCrLf & Err.Description, vbCritical
End Function


