Set args = WScript.Arguments
If args.Count = 0 Then
  WScript.Echo "Workbook path argument is required."
  WScript.Quit 1
End If
wbPath = args(0)
Set fso = CreateObject("Scripting.FileSystemObject")
lockPath = fso.GetParentFolderName(wbPath) & "\\update.lock"
If fso.FileExists(lockPath) Then
  WScript.Echo "Update already in progress."
  WScript.Quit 1
End If
Set lockFile = fso.CreateTextFile(lockPath, True)
lockFile.Close
moduleBase = "https://halyardinc-my.sharepoint.com/:u:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MEL/MODULES/"
objectBase = "https://halyardinc-my.sharepoint.com/:u:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MEL/MICROSOFT_EXCEL_OBJECTS/"
tempPath = fso.BuildPath(fso.GetSpecialFolder(2), "vba_update")
If Not fso.FolderExists(tempPath) Then fso.CreateFolder(tempPath)
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
For Each vbComp In wb.VBProject.VBComponents
  Select Case vbComp.Type
    Case 1
      compName = vbComp.Name
      url = moduleBase & compName & ".bas"
      tempFile = fso.BuildPath(tempPath, compName & ".bas")
      If Download(url, tempFile) Then
        wb.VBProject.VBComponents.Remove vbComp
        Set vbComp = wb.VBProject.VBComponents.Import(tempFile)
        vbComp.Name = compName
      End If
    Case 100
      url = objectBase & vbComp.Name & ".cls"
      tempFile = fso.BuildPath(tempPath, vbComp.Name & ".cls")
      If Download(url, tempFile) Then
        If vbComp.CodeModule.CountOfLines > 0 Then
          vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
        End If
        vbComp.CodeModule.AddFromFile tempFile
      End If
  End Select
Next
wb.Save
wb.Close
xl.AutomationSecurity = 1
Set wb = xl.Workbooks.Open(wbPath)
xl.Visible = True
On Error Resume Next
xl.Run "'" & wb.Name & "'!ShowUpdateSuccess"
On Error GoTo 0
Set wb = Nothing
Set xl = Nothing
If fso.FileExists(lockPath) Then fso.DeleteFile lockPath
WScript.Quit 0

Sub CleanUp()
  If Not wb Is Nothing Then wb.Close False
  If Not xl Is Nothing Then xl.Quit
  If Not fso Is Nothing Then
    If lockPath <> "" Then
      If fso.FileExists(lockPath) Then fso.DeleteFile lockPath
    End If
  End If
End Sub

Function Download(url, path)
  On Error Resume Next
  Dim http, stream
  Set http = CreateObject("MSXML2.XMLHTTP")
  http.Open "GET", url, False
  http.send
  If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile path, 2
    stream.Close
    Download = True
  Else
    Download = False
  End If
  On Error GoTo 0
End Function
