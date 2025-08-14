If args.Count > 1 Then vbapwd = args(1)

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
If vbp.Protection <> 0 Then
  On Error Resume Next
  vbp.Unprotect vbapwd
  If Err.Number <> 0 Then
    WScript.Echo "Failed to unlock VBProject: " & Err.Description
    CleanUp
    WScript.Quit 1
  End If
  On Error GoTo 0
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