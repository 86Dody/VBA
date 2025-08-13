Public Const UPDATE_MESSAGE As String = "New version installed. See release notes for details."
Public VBApswd As String

' Launch an external VBScript that updates the VBA project while Excel is closed.
' This avoids modifying code in a running project which would otherwise halt execution.
Sub updates()

    Dim scriptPath As String
    Dim fnum As Integer
    Dim script As String

    scriptPath = Environ("TEMP") & "\vba_update.vbs"

    script = "Set xl = CreateObject(""Excel.Application"")" & vbCrLf & _
             "xl.Visible = False" & vbCrLf & _
             "xl.AutomationSecurity = 3" & vbCrLf & _
             "Set wb = xl.Workbooks.Open(""" & ThisWorkbook.FullName & """)" & vbCrLf & _
             "Set vbp = wb.VBProject" & vbCrLf & _
             "If vbp.Protection <> 0 Then" & vbCrLf & _
             " xl.VBE.MainWindow.Visible = True" & vbCrLf & _
             " vbp.VBE.CommandBars(""Menu Bar"").Controls(""Tools"").Controls(""VBAProject Properties..."").Execute" & vbCrLf & _
             " xl.SendKeys """ & VBApswd & """ & Chr(13), True" & vbCrLf & _
             " xl.VBE.MainWindow.Visible = False" & vbCrLf & _
             "End If" & vbCrLf & _
             "moduleURL = ""https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MODULES/""" & vbCrLf & _
             "objectURL = ""https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MICROSOFT_EXCEL_OBJECTS/""" & vbCrLf & _
             "tempFolder = CreateObject(""WScript.Shell"").ExpandEnvironmentStrings(""%TEMP%"") & ""\""" & vbCrLf & _
             "For Each vbComp In wb.VBProject.VBComponents" & vbCrLf & _
             " fileURL = """"" & vbCrLf & _
             " tmpFile = """"" & vbCrLf & _
             " Select Case vbComp.Type" & vbCrLf & _
             "  Case 1" & vbCrLf & _
             "   If vbComp.Name <> ""m_update"" Then" & vbCrLf & _
             "    compName = vbComp.Name" & vbCrLf & _
             "    fileURL = moduleURL & compName & "".bas""" & vbCrLf & _
             "    tmpFile = tempFolder & compName & "".bas""" & vbCrLf & _
             "    If DownloadFile(fileURL, tmpFile) Then" & vbCrLf & _
             "     wb.VBProject.VBComponents.Remove vbComp" & vbCrLf & _
             "     Set vbComp = wb.VBProject.VBComponents.Import(tmpFile)" & vbCrLf & _
             "     vbComp.Name = compName" & vbCrLf & _
             "    End If" & vbCrLf & _
             "   End If" & vbCrLf & _
             "  Case 100" & vbCrLf & _
             "   fileURL = objectURL & vbComp.Name & "".cls""" & vbCrLf & _
             "   tmpFile = tempFolder & vbComp.Name & "".cls""" & vbCrLf & _
             "   If DownloadFile(fileURL, tmpFile) Then" & vbCrLf & _
             "     If vbComp.CodeModule.CountOfLines > 0 Then vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines" & vbCrLf & _
             "     vbComp.CodeModule.AddFromFile tmpFile" & vbCrLf & _
             "   End If" & vbCrLf & _
             " End Select" & vbCrLf & _
             "Next" & vbCrLf & _
             "wb.Save" & vbCrLf & _
             "wb.Close False" & vbCrLf & _
             "xl.Quit" & vbCrLf & _
             "Function DownloadFile(url, dest)" & vbCrLf & _
             " On Error Resume Next" & vbCrLf & _
             " Set http = CreateObject(""MSXML2.XMLHTTP"")" & vbCrLf & _
             " http.Open ""GET"", url, False" & vbCrLf & _
             " http.send" & vbCrLf & _
             " If http.Status = 200 Then" & vbCrLf & _
             "  Set stream = CreateObject(""ADODB.Stream"")" & vbCrLf & _
             "  stream.Type = 1" & vbCrLf & _
             "  stream.Open" & vbCrLf & _
             "  stream.Write http.responseBody" & vbCrLf & _
             "  stream.SaveToFile dest, 2" & vbCrLf & _
             "  stream.Close" & vbCrLf & _
             "  DownloadFile = True" & vbCrLf & _
             " Else" & vbCrLf & _
             "  DownloadFile = False" & vbCrLf & _
             " End If" & vbCrLf & _
             "End Function"

    fnum = FreeFile
    Open scriptPath For Output As #fnum
    Print #fnum, script
    Close #fnum

    Shell "wscript """ & scriptPath & """", vbHide
    ThisWorkbook.Close SaveChanges:=False

End Sub
