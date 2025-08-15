Public latestVersion As Long
Public updateMessage As String

Private Const VERSION_URL As String = _
    "https://halyardinc-my.sharepoint.com/:u:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MEL/latest_version.txt"

Private Const LOCK_FILE As String = "update.lock"
Private Const SCRIPT_URL As String = _
    "https://halyardinc-my.sharepoint.com/:u:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MEL/vba_update.vbs"

Public Function GetLatestVersion() As Long
    On Error GoTo errHandler
    Dim http As Object
    Dim content As String
    Dim lines() As String
    Dim i As Long
    Dim startIdx As Long
    Dim endIdx As Long
    Dim msg As String

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", VERSION_URL, False
    http.send
    If http.Status = 200 Then
        content = Replace(http.responseText, vbCr, "")
        lines = Split(content, vbLf)
        If UBound(lines) >= 0 Then
            GetLatestVersion = CLng(lines(0))
        End If
        For i = 1 To UBound(lines)
            If lines(i) = "##UPDATES:##" Then startIdx = i + 1
            If lines(i) = "##END##" Then
                endIdx = i - 1
                Exit For
            End If
        Next i
        If startIdx > 0 And endIdx >= startIdx Then
            For i = startIdx To endIdx
                msg = msg & lines(i) & vbCrLf
            Next i
            updateMessage = Trim(msg)
        End If
    End If
    Exit Function
errHandler:
    GetLatestVersion = 0
    updateMessage = ""
End Function

' Launch an external VBScript that updates the VBA project while Excel is closed.
' This avoids modifying code in a running project which would otherwise halt execution.
Sub updates()


    updating = True
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\vba_update.vbs"

    Dim http As Object, stream As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", SCRIPT_URL, False
    http.send
    If http.Status <> 200 Then
        MsgBox "Update script not found: " & SCRIPT_URL & vbCrLf & _
               "Please contact abel@halyard.ca", vbCritical
        ThisWorkbook.Close SaveChanges:=False
        Application.Quit
    End If
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile scriptPath, 2
    stream.Close

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
    MsgBox updateMessage, vbInformation
    SaveSetting "MELWorkbook", "Update", "ShownVersion", latestVersion
End Sub

Public Function IsUpdatingInProgress() As Boolean
    Dim lockPath As String
    lockPath = ThisWorkbook.Path & "\" & LOCK_FILE
    IsUpdatingInProgress = (Dir(lockPath) <> "")
End Function

Public Function IsOnlyUser() As Boolean
    Dim users As Variant
    Dim count As Long
    On Error Resume Next
    users = ThisWorkbook.UserStatus
    On Error GoTo 0
    If IsArray(users) Then
        count = UBound(users) - LBound(users) + 1
        IsOnlyUser = (count <= 1)
    Else
        IsOnlyUser = True
    End If
End Function

Public Sub ShowUpdateInfoOnce()
    Dim seenVersion As Long
    seenVersion = CLng(GetSetting("MELWorkbook", "Update", "ShownVersion", 0))
    If latestVersion > seenVersion And updateMessage <> "" Then
        MsgBox updateMessage, vbInformation
        SaveSetting "MELWorkbook", "Update", "ShownVersion", latestVersion
    End If
End Sub
