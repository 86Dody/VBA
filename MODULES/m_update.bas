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
    
    script = _
    "Set xl=CreateObject(" & q & "Excel.Application" & q & "):xl.Visible=False:xl.AutomationSecurity=3:" & _
    "Set wb=xl.Workbooks.Open(" & q & ThisWorkbook.FullName & q & "):Set vbp=wb.VBProject:" & _
    "If vbp.Protection<>0 Then xl.VBE.MainWindow.Visible=True:vbp.VBE.CommandBars(" & q & "Menu Bar" & q & ").Controls(" & q & "Tools" & q & ").Controls(" & q & "VBAProject Properties..." & q & ").Execute:" & _
    "xl.SendKeys " & q & VBApswd & q & " & Chr(13),True:xl.VBE.MainWindow.Visible=False:End If:" & _
    "moduleURL=" & q & "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MODULES/" & q & ":" & _
    "objectURL=" & q & "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MICROSOFT_EXCEL_OBJECTS/" & q & ":" & _
    "tempFolder=CreateObject(" & q & "WScript.Shell" & q & ").ExpandEnvironmentStrings(" & q & "%TEMP%" & q & ") & " & q & "\" & q & ":" & _
    "For Each vbComp In wb.VBProject.VBComponents:fileURL=" & q & q & ":tmpFile=" & q & q & ":Select Case vbComp.Type:" & _
    "Case 1:If vbComp.Name<>" & q & "m_update" & q & " Then:compName=vbComp.Name:fileURL=moduleURL & compName & " & q & ".bas" & q & ":" & _
    "tmpFile=tempFolder & compName & " & q & ".bas" & q & ":If DownloadFile(fileURL,tmpFile) Then:wb.VBProject.VBComponents.Remove vbComp:" & _
    "Set vbComp=wb.VBProject.VBComponents.Import(tmpFile):vbComp.Name=compName:End If:End If:" & _
    "Case 100:fileURL=objectURL & vbComp.Name & " & q & ".cls" & q & ":tmpFile=tempFolder & vbComp.Name & " & q & ".cls" & q & ":" & _
    "If DownloadFile(fileURL,tmpFile) Then:If vbComp.CodeModule.CountOfLines>0 Then vbComp.CodeModule.DeleteLines 1,vbComp.CodeModule.CountOfLines:" & _
    "vbComp.CodeModule.AddFromFile tmpFile:End If:" & _
    "End Select:Next:wb.Save:wb.Close False:xl.Quit:" & _
    "Function DownloadFile(url,dest):On Error Resume Next:Set http=CreateObject(" & q & "MSXML2.XMLHTTP" & q & "):" & _
    "http.Open " & q & "GET" & q & ",url,False:http.send:If http.Status=200 Then:Set stream=CreateObject(" & q & "ADODB.Stream" & q & "):" & _
    "stream.Type=1:stream.Open:stream.Write http.responseBody:stream.SaveToFile dest,2:stream.Close:DownloadFile=True:" & _
    "Else:DownloadFile=False:End If:End Function"
 
    fnum = FreeFile
    Open scriptPath For Output As #fnum
    Print #fnum, script
    Close #fnum

    Shell "wscript """ & scriptPath & """", vbHide
    ThisWorkbook.Close SaveChanges:=False

End Sub

