Set args = WScript.Arguments
If args.Count = 0 Then
  WScript.Echo "Workbook path argument is required."
  WScript.Quit 1
End If
wbPath = args(0)
On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
  WScript.Echo "Could not start Excel: " & Err.Description
  CleanUp
  WScript.Quit 1
End If
xl.Visible = False
xl.AutomationSecurity = 3
Set wb = xl.Workbooks.Open(wbPath)
If Err.Number <> 0 Then
  WScript.Echo "Could not open workbook: " & Err.Description
  CleanUp
  WScript.Quit 1
End If
Set vbp = wb.VBProject
If Err.Number <> 0 Then
  WScript.Echo "Could not access VBProject: " & Err.Description
  CleanUp
  WScript.Quit 1
End If
moduleURL="https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MODULES/"
objectURL="https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MICROSOFT_EXCEL_OBJECTS/"
tempFolder=CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\\"
For Each vbComp In wb.VBProject.VBComponents
  fileURL=""
  tmpFile=""
  Select Case vbComp.Type
    Case 1
      If vbComp.Name<>"m_update" Then
        compName=vbComp.Name
        fileURL=moduleURL & compName & ".bas"
        tmpFile=tempFolder & compName & ".bas"
        If DownloadFile(fileURL,tmpFile) Then
          wb.VBProject.VBComponents.Remove vbComp
          Set vbComp=wb.VBProject.VBComponents.Import(tmpFile)
          vbComp.Name=compName
        End If
      End If
    Case 100
      fileURL=objectURL & vbComp.Name & ".cls"
      tmpFile=tempFolder & vbComp.Name & ".cls"
      If DownloadFile(fileURL,tmpFile) Then
        If vbComp.CodeModule.CountOfLines>0 Then
          vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
        End If
        vbComp.CodeModule.AddFromFile tmpFile
      End If
  End Select
Next
wb.Save
CleanUp
If Err.Number <> 0 Then WScript.Quit 1
Sub CleanUp()
  If Not wb Is Nothing Then wb.Close False
  If Not xl Is Nothing Then xl.Quit
End Sub
Function DownloadFile(url,dest)
  On Error Resume Next
  Set http=CreateObject("MSXML2.XMLHTTP")
  http.Open "GET",url,False
  http.send
  If http.Status=200 Then
    Set stream=CreateObject("ADODB.Stream")
    stream.Type=1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile dest,2
    stream.Close
    DownloadFile=True
  Else
    DownloadFile=False
  End If
End Function
