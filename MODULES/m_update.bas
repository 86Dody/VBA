Public Const UPDATE_MESSAGE As String = "New version installed. See release notes for details."
Public VBApswd As String

' Launch an external VBScript that updates the VBA project while Excel is closed.
' This avoids modifying code in a running project which would otherwise halt execution.
Sub updates()

    Dim scriptPath As String
    Dim fnum As Integer
    Dim script As String

    scriptPath = Environ("TEMP") & "\vba_update.vbs"

    Dim q As String
    q = Chr$(34)
    
    Dim lines As Variant
    lines = Array( _
        "Set xl=CreateObject(" & q & "Excel.Application" & q & ")", _
        "xl.Visible=False", _
        "xl.AutomationSecurity=3", _
        "Set wb=xl.Workbooks.Open(" & q & ThisWorkbook.FullName & q & ")", _
        "Set vbp=wb.VBProject", _
        "If vbp.Protection<>0 Then", _
        "  xl.VBE.MainWindow.Visible=True", _
        "  vbp.VBE.CommandBars(" & q & "Menu Bar" & q & ").Controls(" & q & "Tools" & q & ").Controls(" & q & "VBAProject Properties..." & q & ").Execute", _
        "  xl.SendKeys " & q & VBApswd & q & " & Chr(13),True", _
        "  xl.VBE.MainWindow.Visible=False", _
        "End If", _
        "moduleURL=" & q & "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MODULES/" & q, _
        "objectURL=" & q & "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MICROSOFT_EXCEL_OBJECTS/" & q, _
        "tempFolder=CreateObject(" & q & "WScript.Shell" & q & ").ExpandEnvironmentStrings(" & q & "%TEMP%" & q & ") & " & q & "\" & q, _
        "For Each vbComp In wb.VBProject.VBComponents", _
        "  fileURL=" & q & q, _
        "  tmpFile=" & q & q, _
        "  Select Case vbComp.Type", _
        "    Case 1", _
        "      If vbComp.Name<>" & q & "m_update" & q & " Then", _
        "        compName=vbComp.Name", _
        "        fileURL=moduleURL & compName & " & q & ".bas" & q, _
        "        tmpFile=tempFolder & compName & " & q & ".bas" & q, _
        "        If DownloadFile(fileURL,tmpFile) Then", _
        "          wb.VBProject.VBComponents.Remove vbComp", _
        "          Set vbComp=wb.VBProject.VBComponents.Import(tmpFile)", _
        "          vbComp.Name=compName", _
        "        End If", _
        "      End If", _
        "    Case 100", _
        "      fileURL=objectURL & vbComp.Name & " & q & ".cls" & q, _
        "      tmpFile=tempFolder & vbComp.Name & " & q & ".cls" & q, _
        "      If DownloadFile(fileURL,tmpFile) Then", _
        "        If vbComp.CodeModule.CountOfLines>0 Then", _
        "          vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines", _
        "        End If", _
        "        vbComp.CodeModule.AddFromFile tmpFile", _
        "      End If", _
        "  End Select", _
        "Next", _
        "wb.Save", _
        "wb.Close False", _
        "xl.Quit", _
        "Function DownloadFile(url,dest)", _
        "  On Error Resume Next", _
        "  Set http=CreateObject(" & q & "MSXML2.XMLHTTP" & q & ")", _
        "  http.Open " & q & "GET" & q & ",url,False", _
        "  http.send", _
        "  If http.Status=200 Then", _
        "    Set stream=CreateObject(" & q & "ADODB.Stream" & q & ")", _
        "    stream.Type=1", _
        "    stream.Open", _
        "    stream.Write http.responseBody", _
        "    stream.SaveToFile dest,2", _
        "    stream.Close", _
        "    DownloadFile=True", _
        "  Else", _
        "    DownloadFile=False", _
        "  End If", _
        "End Function" _
    )

    script = Join(lines, vbCrLf)
 
    fnum = FreeFile
    Open scriptPath For Output As #fnum
    Print #fnum, script
    Close #fnum

    If Dir(scriptPath) = "" Then
        MsgBox "Update script was not created: " & scriptPath, vbExclamation
        Exit Sub
    End If

    On Error GoTo ShellError
    Shell """" & Environ("WINDIR") & "\System32\wscript.exe"" """" & scriptPath & """"", vbHide
    On Error GoTo 0
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub

ShellError:
    MsgBox "Failed to run update script: " & Err.Description, vbCritical

End Sub

