Attribute VB_Name = "m_expandMEL"
Sub expand_MEL()

On Error GoTo error3

    Application.ScreenUpdating = False
    Application.EnableEvents = False
      
    Set network = CreateObject("WScript.Network")
      
    Dim ws As Worksheet
    Set ws = Sheets("MEL")
    
If Sheets("MEL").Buttons("button 39").Enabled = True And access < 3 And WorksheetFunction.CountIf(Range("MEL_LST[EQUIPMENT DESCRIPTION]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TAG]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[WBS]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TYPE]"), "") = 0 Then
   
    
    Sheets("MEL").Unprotect Password:=pswd
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("MEL_LST")

    tbl.ListRows.Add
    
    Sheets("MEL").Cells(Range("MEL_ROWS").Value + 6, Range("MEL_LST[[#Headers],[NUMBER]]").Column).Value = Range("MEL_ROWS").Value
    
    If Sheets("MEL").Range("Version").Value = "START" Then
    
        Sheets("MEL").Cells(Range("MEL_LST").Rows.Count + 6, Range("MEL_LST[[#Headers],[REV]]").Column).Value = "A"
    
    Else
    
        Sheets("MEL").Cells(Range("MEL_ROWS").Value + 6, Range("MEL_LST[[#Headers],[REV]]").Column).Value = Sheets("MEL").Range("VERSION").Value
    
    End If
    
    Sheets("MEL").Cells(Range("MEL_ROWS").Value + 6, Range("MEL_LST[[#Headers],[DATE]]").Column).Value = Format(Date, "yyyy/mm/dd")
    Sheets("MEL").Cells(Range("MEL_ROWS").Value + 6, Range("MEL_LST[[#Headers],[CONTROL]]").Column).Value = network.UserName
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If Range("A3") = "M" Then

    Else
        Sheets("MEL").Range("B:AJ").EntireColumn.Hidden = False
        Sheets("MEL").Protect Password:=pswd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Userinterfaceonly:=True, AllowFormattingColumns:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
    End If
    
    Sheets("MEL").Range("MEL_LST[[#Headers],[NUMBER]]").End(xlDown).Select
        
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Set ws = Nothing
    Set tbl = Nothing
    Set network = Nothing

Call cellBlock

ElseIf access = 3 And Sheets("MEL").Buttons("button 39").Enabled = True And WorksheetFunction.CountIf(Range("MEL_LST[EQUIPMENT DESCRIPTION]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TAG]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[WBS]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TYPE]"), "") = 0 Then

    MsgBox ("You don't have the rights to add a new equipment, please check with the process department")

ElseIf access = 3 And Sheets("MEL").Buttons("button 39").Enabled = False Then

    MsgBox ("Function is temporarily not available." & vbNewLine & _
        "Please contact abel@halyard.ca for more information.")

Else

    MsgBox ("Please complete the previous entries before adding a new row")
   
    
End If

ActiveWorkbook.Save

Exit Sub

error3:    MsgBox ("error3: Procedure - adding a row to the MEL has failed")
            Call start
            Set ws = Nothing
            Set tbl = Nothing
            Set network = Nothing
            Call cellBlock
            Application.ScreenUpdating = True
            Application.EnableEvents = True
    
End Sub




