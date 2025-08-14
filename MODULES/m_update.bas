Public Const UPDATE_MESSAGE As String = "New version installed. See release notes for details."
Public VBApswd As String
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

    Dim scriptPath As String
    scriptPath = "C:\Users\Abel\OneDrive - Halyard Inc\Documents\Abel\Programing\GitHub\VBA\vba_update.vbs"

    If Dir(scriptPath) = "" Then
        MsgBox "Update script not found: " & scriptPath, vbExclamation
        Exit Sub
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
          """" & ThisWorkbook.FullName & """ " & _
          """" & VBApswd & """"

    On Error GoTo ShellError
    Shell cmd, vbHide
    On Error GoTo 0
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub

ShellError:
    MsgBox "Failed to run update script: " & Err.Description, vbCritical

End Sub
