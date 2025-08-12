VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_BeforeClose(Cancel As Boolean)

On Error GoTo error100

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

'save the standard protected range
Call SelectLockedCells

'initiate
Call start

'Call updates

End Sub







