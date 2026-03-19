Attribute VB_Name = "Module1"
Option Explicit

' ===== SharePointPath_v1.bas =====
' Purpose:
'   1) Detect the currently logged-in user's profile path (e.g., C:\Users\Legha).
'   2) Construct the full path to the SharePoint sync folder.
'   3) Place this path into cell G5 and remove old Dropdown validation.
'
' Usage:
'   Run 'SetupSharePointPath_G5' once to configure the sheet.

' ADJUST THIS CONSTANT if the folder name changes in the future
Private Const COMPANY_FOLDER As String = "\BERADER PROPERTIES (PTY) LTD\Berader Properties - Berader Server\"

Public Sub SetupSharePointPath_G5()
    Dim wsTarget As Worksheet
    Dim userProfile As String
    Dim fullPath As String
    
    Set wsTarget = ActiveSheet
    
    ' 1. Get User Profile (e.g., "C:\Users\Legha")
    userProfile = Environ("UserProfile")
    
    ' 2. Build the full path
    fullPath = userProfile & COMPANY_FOLDER
    
    ' 3. Verify the folder actually exists locally
    If Dir(fullPath, vbDirectory) = vbNullString Then
        MsgBox "Warning: The expected SharePoint folder was not found locally:" & vbCrLf & vbCrLf & _
               fullPath & vbCrLf & vbCrLf & _
               "Please ensure OneDrive is running and the folder is synced.", vbExclamation
    End If
    
    ' 4. Update Cell G5
    With wsTarget.Range("G5")
        ' Clear old "Data Validation" dropdowns from the previous version
        .Validation.Delete
        
        ' Set the new value
        .Value = fullPath
    End With
    
    ' Optional: Clear the old config sheet if you want to clean up
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("__Config").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    MsgBox "Path updated to: " & vbCrLf & fullPath, vbInformation
End Sub

Public Sub RefreshSharePointPath()
    ' Helper to simply re-run the setup if needed
    SetupSharePointPath_G5
End Sub

