Public Const UPDATE_MESSAGE As String = "New version installed. See release notes for details."
Public Const LATEST_VERSION As Long = 1
Public VBApswd As String

' Launch an external VBScript that updates the VBA project while Excel is closed.
' This avoids modifying code in a running project which would otherwise halt execution.
Sub updates()

    updating = True

    Dim scriptPath As String
    scriptPath = ThisWorkbook.Path & "\vba_update.vbs"

    If Dir(scriptPath) = "" Then
        MsgBox "Update script not found: " & scriptPath, vbExclamation
        Exit Sub
    End If

    ' record the version installed
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="VbaVersion", RefersTo:="=" & LATEST_VERSION
    If Err.Number <> 0 Then
        ThisWorkbook.Names("VbaVersion").RefersTo = "=" & LATEST_VERSION
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
