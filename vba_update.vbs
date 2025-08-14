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
modulePath="C:\Users\Abel\OneDrive - Halyard Inc\Documents\Abel\Programing\GitHub\VBA\MODULES\"
objectPath="C:\Users\Abel\OneDrive - Halyard Inc\Documents\Abel\Programing\GitHub\VBA\MICROSOFT_EXCEL_OBJECTS\"
Set fso=CreateObject("Scripting.FileSystemObject")
For Each vbComp In wb.VBProject.VBComponents
  Select Case vbComp.Type
    Case 1
      compName=vbComp.Name
      filePath=modulePath & compName & ".bas"
      If fso.FileExists(filePath) Then
        wb.VBProject.VBComponents.Remove vbComp
        Set vbComp=wb.VBProject.VBComponents.Import(filePath)
        vbComp.Name=compName
      End If
    Case 100
      filePath=objectPath & vbComp.Name & ".cls"
      If fso.FileExists(filePath) Then
        If vbComp.CodeModule.CountOfLines>0 Then
          vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
        End If
        vbComp.CodeModule.AddFromFile filePath
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
WScript.Quit 0

Sub CleanUp()
  If Not wb Is Nothing Then wb.Close False
  If Not xl Is Nothing Then xl.Quit
End Sub
