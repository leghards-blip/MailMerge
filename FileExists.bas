Attribute VB_Name = "FileExists"
Function FileCheck(filePath As Variant) As String
    On Error GoTo HandleError

    If IsError(filePath) Then
        FileCheck = "? Invalid input"
    ElseIf Trim(filePath) = "" Then
        FileCheck = "? Missing path"
    ElseIf Dir(CStr(filePath)) <> "" Then
        FileCheck = "?? Found"
    Else
        FileCheck = "? Not Found"
    End If
    Exit Function

HandleError:
    FileCheck = "? Error"
End Function

