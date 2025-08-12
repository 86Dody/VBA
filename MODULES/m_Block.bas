Attribute VB_Name = "m_Block"
Sub cellBlock()

On Error GoTo error9

Dim wb As Workbook
Dim ws As Worksheet

Set wb = Workbooks("M25058-2000-M-MEL-003.xlsm")
Set ws = wb.Sheets("MEL")

If ws.Range("MEL_LST[TYPE]").Rows.Count > 2 And ws.Range("VERSION").Value <> "START" Then
    
    ws.Range("MEL_LST[TYPE]").Locked = True
    ws.Cells(ws.Range("MEL_LST[TYPE]").Rows.Count + 6, ws.Range("MEL_LST[TYPE]").Column).Locked = False
    ws.Cells(ws.Range("MEL_LST[TYPE]").Rows.Count + 6 - 1, ws.Range("MEL_LST[TYPE]").Column).Locked = False

Else

    ws.Range("MEL_LST[TYPE]").Locked = False

End If

Set ws = Nothing
Set wb = Nothing


Exit Sub

error9: MsgBox ("error9: Procedure - blocking TYPE has failed")

End Sub
