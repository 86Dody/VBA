Public Const UPDATE_MESSAGE As String = _
    "Update successful. This release includes the latest features and bug fixes."
Public latestVersion As Long

Private Const VERSION_URL As String = _
    "https://halyardinc-my.sharepoint.com/:u:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/latest_version.txt"

Public Function GetLatestVersion() As Long
    On Error GoTo errHandler
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", VERSION_URL, False
    http.send
    If http.Status = 200 Then
        GetLatestVersion = CLng(http.responseText)
    End If
    Exit Function
errHandler:
    GetLatestVersion = 0
End Function

' Launch an external VBScript that updates the VBA project while Excel is closed.
' This avoids modifying code in a running project which would otherwise halt execution.
Sub updates()


    updating = True

    MsgBox "The application will now update and reopen with the new version.", _
           vbInformation

    Dim scriptPath As String
    scriptPath = "C:\Users\Abel\OneDrive - Halyard Inc\Documents\Abel\Programing\GitHub\VBA\vba_update.vbs"

    If Dir(scriptPath) = "" Then
        MsgBox "Update script not found: " & scriptPath & vbCrLf & _
               "Please contact abel@halyard.ca", vbCritical
        ThisWorkbook.Close SaveChanges:=False
        Application.Quit
    End If

    ' record the version installed
    If latestVersion = 0 Then latestVersion = GetLatestVersion()
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="VbaVersion", RefersTo:="=" & latestVersion
    If Err.Number <> 0 Then
        ThisWorkbook.Names("VbaVersion").RefersTo = "=" & latestVersion
        Err.Clear
    End If
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    cmd = """" & Environ("WINDIR") & "\System32\wscript.exe"" " & _
          """" & scriptPath & """ " & _
          """" & ThisWorkbook.FullName & """"

    On Error GoTo ShellError
    Shell cmd, vbHide
    On Error GoTo 0
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub

ShellError:
    MsgBox "Update failed. Please contact abel@halyard.ca", vbCritical
    ThisWorkbook.Close SaveChanges:=False
    Application.Quit

End Sub

Sub ShowUpdateSuccess()
    MsgBox UPDATE_MESSAGE, vbInformation
End Sub
