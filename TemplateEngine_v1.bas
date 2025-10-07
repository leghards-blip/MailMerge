
Attribute VB_Name = "TemplateEngine_v1"
Option Explicit

' ===== TemplateEngine_v1.bas =====
' Load an email body template from .txt/.html or .docx and replace placeholders.
' Placeholders use double curly braces, e.g. {{UNIT}}, {{COMPLEX}}, {{MONTHYEAR}}, {{OWNERNAME}}
'
' Public:
'   LoadTemplateText(path) As String                ' .txt or .html as-is
'   LoadDocxAsHtml(path) As String                  ' via Word automation -> filtered HTML
'   RenderTemplate(templateText, keys, values)      ' replace {{KEY}} with value (arrays must align)
'   BuildEmailHtmlFromFile(path, keys, values)      ' auto-detect by extension; returns HTML string
'
' Notes:
'   - For .txt files, line breaks will be converted to <br> tags for HTML email.
'   - For .docx, Word must be installed; uses SaveAs2 with wdFormatFilteredHTML.
'   - For security and portability, prefer .html or .txt templates when possible.
'   - Keys are case-insensitive ("UNIT" == "unit").
'
' Example usage:
'   Dim html As String
'   html = BuildEmailHtmlFromFile("Z:\Templates\levy_email.docx", _
'           Array("UNIT","COMPLEX","MONTHYEAR"), _
'           Array(unitNo, complexCode, monthYear))
'
Private Function ReadFileText(ByVal filePath As String) As String
    Dim f As Integer, s As String
    f = FreeFile(0)
    Open filePath For Binary As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    ReadFileText = s
End Function

Public Function LoadTemplateText(ByVal filePath As String) As String
    ' Returns text/HTML as-is from .txt or .html files
    LoadTemplateText = ReadFileText(filePath)
End Function

Public Function LoadDocxAsHtml(ByVal filePath As String) As String
    ' Opens a .docx in Word and returns filtered HTML as a string
    Dim wApp As Object, wDoc As Object
    Dim tempHtml As String, html As String
    Dim wdFormatFilteredHTML As Long: wdFormatFilteredHTML = 10
    
    tempHtml = Environ$("TEMP") & "\" & "template_" & Format(Now, "yyyymmdd_hhnnss") & ".htm"
    
    On Error GoTo Fail
    Set wApp = CreateObject("Word.Application")
    wApp.Visible = False
    Set wDoc = wApp.Documents.Open(filePath, ReadOnly:=True)
    wDoc.SaveAs2 FileName:=tempHtml, FileFormat:=wdFormatFilteredHTML
    wDoc.Close False
    wApp.Quit
    
    html = ReadFileText(tempHtml)
    On Error Resume Next
    Kill tempHtml
    On Error GoTo 0
    
    LoadDocxAsHtml = html
    Exit Function
Fail:
    On Error Resume Next
    If Not wDoc Is Nothing Then wDoc.Close False
    If Not wApp Is Nothing Then wApp.Quit
    On Error GoTo 0
    Err.Raise vbObjectError + 513, "TemplateEngine_v1.LoadDocxAsHtml", "Failed to load DOCX via Word. " & Err.Description
End Function

Public Function RenderTemplate(ByVal templateText As String, ByVal keys As Variant, ByVal values As Variant) As String
    ' Replace {{KEY}} tokens (case-insensitive) with values.
    Dim i As Long, token As String, result As String
    result = templateText
    For i = LBound(keys) To UBound(keys)
        token = "{{" & UCase$(CStr(keys(i))) & "}}"
        result = ReplaceCI(result, token, CStr(values(i)))
    Next i
    RenderTemplate = result
End Function

Public Function BuildEmailHtmlFromFile(ByVal filePath As String, ByVal keys As Variant, ByVal values As Variant) As String
    Dim ext As String, raw As String, html As String
    
    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
    Select Case ext
        Case "txt"
            raw = LoadTemplateText(filePath)
            raw = HtmlEncode(raw)
            raw = Replace(raw, vbCrLf, "<br>")
            raw = Replace(raw, vbLf, "<br>")
            html = RenderTemplate(raw, keys, values)
        Case "htm", "html"
            raw = LoadTemplateText(filePath)
            html = RenderTemplate(raw, keys, values)
        Case "docx"
            raw = LoadDocxAsHtml(filePath)
            html = RenderTemplate(raw, keys, values)
        Case Else
            Err.Raise vbObjectError + 514, "TemplateEngine_v1.BuildEmailHtmlFromFile", "Unsupported template type: " & ext
    End Select
    
    BuildEmailHtmlFromFile = html
End Function

' === Helpers =================================================================

Private Function ReplaceCI(ByVal text As String, ByVal findToken As String, ByVal repl As String) As String
    ' Case-insensitive replace
    Dim pos As Long, res As String, L As Long, fU As String, tU As String
    fU = UCase$(findToken)
    tU = UCase$(text)
    L = Len(findToken)
    pos = InStr(1, tU, fU, vbTextCompare)
    Do While pos > 0
        res = res & Mid$(text, 1, pos - 1) & repl
        text = Mid$(text, pos + L)
        tU = UCase$(text)
        pos = InStr(1, tU, fU, vbTextCompare)
    Loop
    ReplaceCI = res & text
End Function

Private Function HtmlEncode(ByVal s As String) As String
    ' Minimal HTML encoder for <, >, &, "
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    HtmlEncode = s
End Function
