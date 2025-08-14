Private Sub Workbook_BeforeClose(Cancel As Boolean)

On Error GoTo error100

    If updating Then Exit Sub
    If pswd = "" Then pswd = "ConPdeProcess"

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Sheets("MEL").Columns.EntireColumn.Hidden = False
    Sheets("MEL").Range("AL:AV").EntireColumn.Hidden = True
    
    ws.Unprotect Password:=pswd
    ws.Cells.Locked = False
    secure.Locked = True
    
    Sheets("MEL").Protect Password:=pswd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Userinterfaceonly:=True, AllowFormattingColumns:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
    Set ws = Nothing
    
    Exit Sub
    
error100:
    Set ws = Nothing
    Cancel = True
    MsgBox ("error100: Procedure - close sheet has failed, contact Abel if error persists")
    Call start
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Private Sub Workbook_Open()
    Dim versionName As Name
    Dim version As Long

    latestVersion = GetLatestVersion()

    On Error Resume Next
    Set versionName = ThisWorkbook.Names("VbaVersion")
    If Not versionName Is Nothing Then
        version = CLng(Mid(versionName.RefersTo, 2))
    Else
        version = 0
    End If
    On Error GoTo 0

    If latestVersion > 0 And version <> latestVersion Then
        VBApswd = pswd
        updates
        Exit Sub
    End If

    Call SelectLockedCells
    Call start

'HOLA

End Sub











