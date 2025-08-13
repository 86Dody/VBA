
Sub addMotor()

On Error GoTo error4

Application.ScreenUpdating = False
Application.EnableEvents = False

Dim ws As Worksheet
Set ws = Sheets("MEL")

Dim tbl As ListObject
Set tbl = ws.ListObjects("MEL_LST")

Dim Message, Title, Default, MyValue

'loop to add the motors
Dim i As Integer
i = 0

If ws.Buttons("button 38").Enabled = True And access < 3 And WorksheetFunction.CountIf(Range("MEL_LST[EQUIPMENT DESCRIPTION]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TAG]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[WBS]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TYPE]"), "") = 0 Then
    
    Set network = CreateObject("WScript.Network")
    
    'remove filters
    tbl.AutoFilter.ShowAllData
    
    ws.Unprotect Password:=pswd

    'condition: selected column on the table, item must not be a motor already, item must be selectable for a motor
    If Intersect(ActiveCell, tbl.DataBodyRange) Is Nothing Then
    
    MsgBox ("Please select the item you want to add the motor to")
    
    ElseIf ws.Cells(ActiveCell.Row, _
    Range("MEL_LST[[#Headers],[MOTOR]]").Column).Value = "Y" Or WorksheetFunction.XLookup(Sheets("MEL").Cells(ActiveCell.Row, Sheets("MEL").Range("MEL_LST[TYPE]").Column), Sheets("VARIANCES").Range("V_TYPE[TYPE]"), Sheets("VARIANCES").Range("V_TYPE[TYPE_E]")) = "N" Then
        MsgBox ("Error: It is not possible to add a motor to this item")
    Else
    
        
        validation = MsgBox("Do you want to add a motor under " & ws.Cells(ActiveCell.Row, _
        Range("MEL_LST[[#Headers],[TAG]]").Column).Value, vbOKCancel) = vbOKCancel
        
        If validation = False Then
        
            GoTo noTag
            
        End If
        
         
         Message = "How many motors shall be added?"    ' Set prompt.
            Title = "Motor adder"    ' Set title.
            Default = "2"    ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, Default)

            
        
        If MyValue < 2 Then
    
            GoTo noTag1
            
        End If
        
        
        'existing number of motors
        aux = ws.Cells(ActiveCell.Row, Range("MEL_LST[[#Headers],[MOTOR QUANTITY]]").Column).Value
        
        If ws.Cells(ActiveCell.Row, Range("MEL_LST[[#Headers],[MOTOR QUANTITY]]").Column).Value = 1 And ws.Cells(ActiveCell.Row, Range("MEL_LST[[#Headers],[EMERGENCY LOAD]]").Column).Value = "---" Then
            
            aux = aux + 1
            'tbl.ListRows.Add ActiveCell.Row - 6 + aux
        
        ElseIf ws.Cells(ActiveCell.Row, Range("MEL_LST[[#Headers],[EMERGENCY LOAD]]").Column).Value = "-" Then
        
            'tbl.ListRows.Add ActiveCell.Row - 6 + aux
            
        
        Else
        
            'tbl.ListRows.Add ActiveCell.Row - 6 + aux
        
        End If
        
        Do While i < MyValue
        
            tbl.ListRows.Add ActiveCell.Row - 6 + aux + i
        
        If ws.Range("Version").Value = "START" Then
        
            ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[REV]]").Column).Value = "A"
        
            Else
        
            ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[REV]]").Column).Value = ws.Range("VERSION").Value
        
        End If
        
        ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[DATE]]").Column).Value = Format(Date, "yyyy/mm/dd")
        ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[MOTOR]]").Column).Value = "Y"
        ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[WBS]]").Column).Value = _
        ws.Cells(ActiveCell.Row, Range("MEL_LST[[#Headers],[WBS]]").Column).Value
        ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[TYPE]]").Column).Value = _
        ws.Cells(ActiveCell.Row, Range("MEL_LST[[#Headers],[TYPE]]").Column).Value
        ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[SUPPLY PKG]]").Column).Value = _
        ws.Cells(ActiveCell.Row, Range("MEL_LST[[#Headers],[SUPPLY PKG]]").Column).Value
        ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[NUMBER]]").Column).Value = _
        ws.Cells(ActiveCell.Row, Range("MEL_LST[[#Headers],[NUMBER]]").Column).Value
        ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[CONTROL]]").Column).Value = network.UserName
        ws.Cells(ActiveCell.Row + aux + i, Range("MEL_LST[[#Headers],[PFD]]").Column).Value = _
        ws.Cells(ActiveCell.Row, Range("MEL_LST[[#Headers],[PFD]]").Column).Value
        
        i = i + 1
        
        Loop
        
    End If


    If Range("A3") = "M" Then

    Else
        ws.Protect Password:=pswd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Userinterfaceonly:=True, AllowFormattingColumns:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
    
    End If

    ActiveSheet.Range("MEL_LST").AutoFilter Field:=1, Criteria1:="<>DELETED"

Call cellBlock

ElseIf access = 3 And ws.Buttons("button 38").Enabled = True And WorksheetFunction.CountIf(Range("MEL_LST[EQUIPMENT DESCRIPTION]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TAG]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[WBS]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TYPE]"), "") = 0 Then
 
MsgBox ("You don't have the rights to add a new motor, please check with the process department")

ElseIf ws.Buttons("button 38").Enabled = False Then

MsgBox ("Function is temporarily not available." & vbNewLine & _
        "Please contact abel@halyard.ca for more information.")

Else

MsgBox ("Please complete all previous entries before entering a new motor")

End If


Set ws = Nothing
Set tbl = Nothing
Set network = Nothing

Application.ScreenUpdating = True
Application.EnableEvents = True

ActiveWorkbook.Save

Exit Sub
noTag: MsgBox ("No motor has been added")
Exit Sub
noTag1: MsgBox ("This function is only available for multiple motor addition")
Exit Sub
error4:  MsgBox ("error4: Procedure - Adding line motor has failed")

            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Set ws = Nothing
            Set tbl = Nothing
            Set network = Nothing
            Call cellBlock
            Call start
            
End Sub

