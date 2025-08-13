VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim wb As Workbook
Dim ws As Worksheet
Dim network As Object

Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo error2

Set wb = Workbooks("M25058-2000-M-MEL-003.xlsm")
Set ws = wb.Sheets("MEL")



'Worksheets("MEL").Calculate

Application.EnableEvents = False
Application.ScreenUpdating = False

    ws.Unprotect Password:=pswd
    Set network = CreateObject("WScript.Network")
    
    'Programming to control the changes and give specific users the power to reset the formating
    'defining range for change control of MEL info
    Dim keyCells As Range
    Set keyCells = ws.Range("MEL_LST[[EQUIPMENT DESCRIPTION]:[PHASE]]")
    
    'defining range for change control of MEL CONTROL info
    Dim keyCells1 As Range
    Set keyCells1 = ws.Range("MEL_LST[CONTROL]")
    
    'defining range for change control of MEL VERSION info
    Dim keyCells3 As Range
    Set keyCells3 = ws.Range("VERSION")
    
    'defining range for change control of MEL MODE info
    Dim keyCells4 As Range
    Set keyCells4 = ws.Range("MODE")
    
    'defining range for change control under TYPE selected
    Dim keycells5 As Range
    Set keycells5 = ws.Range("MEL_LST[TYPE]")
    
    
    'defining range for change control under DESCRIPTION
    Dim keycells6 As Range
    Set keycells6 = ws.Range("MEL_LST[EQUIPMENT DESCRIPTION]")
    
    
    'defining range for change control under STEP STATUS
    Dim keycells7 As Range
    Set keycells7 = ws.Range("STEP2A,STEP2B,STEP3")

    
 'changes in table
    
    If Not Application.Intersect(keyCells, Target) Is Nothing And Target.Column > 3 And Target.Column < 32 Then
    
        Call change_table(keyCells, Range(Target.Address), network.UserName)
        
    End If
    
    
 'changes in control column
 
    If Not Application.Intersect(keyCells1, Target) Is Nothing And access <= 2 And ws.Range("VERSION").Value <> "START" Then
    
        Call change_control_column(keyCells1, Range(Target.Address))
       
    End If

'changes after revision change

If Not Application.Intersect(keyCells3, Target) Is Nothing And access <= 2 Then
    
    Call change_revision(keyCells3, Range(Target.Address), network.UserName)
        
End If


'changes in application mode
If Not Application.Intersect(keyCells4, Target) Is Nothing And access = 1 Then

    Call change_mode(keyCells4, Range(Target.Address), network.UserName)

ElseIf Not Application.Intersect(keyCells4, Target) Is Nothing And access <= 2 And ws.Range("VERSION").Value = "START" Then
    
    Call change_mode(keyCells4, Range(Target.Address), network.UserName)
    
ElseIf Not Application.Intersect(keyCells4, Target) Is Nothing Then

    ws.Range("MODE").Value = ""
    MsgBox ("You don't have permissions to change the mode, please contact abel@halyard.ca")
    

End If



'changes in TYPE column control
    
If Not Application.Intersect(keycells5, Target) Is Nothing Then
    
    Call type_selection(keycells5, Range(Target.Address), network.UserName)
        
End If

 'changes in DESCRIPTION to force organizing table
    
    If Not Application.Intersect(keycells6, Target) Is Nothing Then
    
        Call change_description(keycells6, Range(Target.Address), network.UserName)
        
    End If


'changes in STEP STATUS column control
    
If Not Application.Intersect(keycells7, Target) Is Nothing Then
    
    Call chage_stepstatus(keycells7, Range(Target.Address), network.UserName)
        
End If


If Sheets("MEL").Range("A3").Value = "" Then

    Sheets("MEL").Protect Password:=pswd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Userinterfaceonly:=True, AllowFormattingColumns:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
  
End If

'Worksheets("MEL").Calculate
Application.EnableEvents = True
Application.ScreenUpdating = True
Set network = Nothing
Set ws = Nothing
Set wb = Nothing

Exit Sub

error2: MsgBox ("error2: Procedure - MEL control change has failed")
        Call start
        Sheets("MEL").Protect Password:=pswd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Userinterfaceonly:=True, AllowFormattingColumns:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Set network = Nothing
        Set ws = Nothing
        Set wb = Nothing
        'Worksheets("MEL").Calculate


End Sub


Function change_table(ByVal Trow As Range, ByVal Krang As Range, net As String)
    
On Error GoTo error4

        Application.EnableEvents = False
        
        If ws.Range("VERSION").Value <> "START" Then
        
        Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[CONTROL]]").Column) = net
        Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[REV]]").Column) = ws.Range("VERSION").Value
        Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[DATE]]").Column) = Format(Date, "yyyy/mm/dd")
        Krang.Interior.ColorIndex = 3
        
        Else
        
        Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[REV]]").Column) = "A"
        Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[DATE]]").Column) = Format(Date, "yyyy/mm/dd")
        
        End If
        
        Application.EnableEvents = True

Exit Function

error4: MsgBox ("error4: Procedure - change in table has failed")
        Application.EnableEvents = True
        Call start

End Function



Function change_control_column(ByVal Trow As Range, ByVal Krang As Range)

On Error GoTo error5

Dim ii As Integer
Dim i As Integer
Dim aux As Integer
Dim aux2 As Integer
Dim tbl As Table

Application.EnableEvents = False

If ws.Range("MODE").Value = "PAUSE" And Krang.Columns.Count >= 46 Then

    GoTo Deleted

End If


If Not ws.Cells(Krang.Row, Range("MEL_LST[CONTROL]").Column).Value = "Deleted" Then

        Dim validation As Boolean
        validation = MsgBox("Are you sure you want to accept the changes?", vbOKCancel, "CHANGE ACCEPTANCE") = vbOK

        
        If validation = True Then
       
        For ii = 1 To ws.Range("MEL_LST").Rows.Count
                
            If ii = Krang.Row - 6 Then
            
                For i = 1 To ws.Range("MEL_LST").Columns.Count
                    
                    If Cells(Krang.Row, i).Interior.ColorIndex = 3 Then
                        
                        Cells(Krang.Row, i).Interior.ColorIndex = 4
                        
                    End If
                    
                Next
            
            End If
        
        Next
                

        
        Else
        
        Krang.Value = "Pending"
        
        End If
        
Else

    'Manage deletion of rows
    'Special cases for Agitator and Thickener only (associated equipments)
    
    aux = ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[MOTOR QUANTITY]]").Column).Value
    aux2 = ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[MOTOR QUANTITY]]").Column).Value
     
    If ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[TYPE]]").Column).Value = "AG" And _
    ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[MOTOR]]").Column).Value = "Y" Then
        
        validation = MsgBox("You have deleted the item " & ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[TAG]]").Column).Value & _
        vbNewLine & "Do you want to delete the associated tank with tag " & ws.Cells(Krang.Row + aux + 1, _
        ws.Range("MEL_LST[[#Headers],[TAG]]").Column).Value) = vbOKCancel
        
        If validation = True Then
            
            For i = 0 To aux + 1
                ws.Cells(Krang.Row + i, ws.Range("MEL_LST[[#Headers],[Control]]").Column).Value = "Deleted"
            Next
            
        Else
            MsgBox ("The tank " & ws.Cells(Krang.Row, _
            ws.Range("MEL_LST[[#Headers],[TAG]]").Column).Value & " associate with the agitator " & _
            ws.Cells(Krang.Row + aux + 1, ws.Range("MEL_LST[[#Headers],[TAG]]").Column).Value & " hasn't been deleted")
        End If
         
    ElseIf ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[TYPE]]").Column).Value = "AG" And _
    ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[MOTOR]]").Column).Value <> "Y" Then
        
        validation = MsgBox("You have deleted the item " & ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[TAG]]").Column).Value & _
        vbNewLine & "Do you want to delete the associated tank with tag " & _
        ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[TAG]]").Column).Value) = vbOKCancel
         
        If validation = True Then
            
            For i = 0 To aux
                ws.Cells(Krang.Row + i, ws.Range("MEL_LST[[#Headers],[Control]]").Column).Value = "Deleted"
            Next
            
        Else
        
        End If
         
         
    ElseIf ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[TYPE]]").Column).Value = "TH" And _
    ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[TYPE]]").Column).Value = "TM" And _
    aux2 > 1 Then
    
        For i = 0 To aux2 + 1
            ws.Cells(Krang.Row + i, ws.Range("MEL_LST[[#Headers],[Control]]").Column).Value = "Deleted"
        Next
    
    ElseIf ws.Cells(Krang.Row + 2, ws.Range("MEL_LST[[#Headers],[MOTOR]]").Column).Value = "Y" And aux2 = 1 Then
        For i = 0 To aux2 + 1
            ws.Cells(Krang.Row + i, ws.Range("MEL_LST[[#Headers],[Control]]").Column).Value = "Deleted"
        Next
    
    ElseIf ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[TYPE]]").Column).Value <> "TM" Then
    
        MsgBox ("ATTENTION: The thickener mechanism associated -TM- can not be located." & vbNewLine & _
        "Deletion process of the mechanism and the associated motors must be completed manually.")
    
    Else
        For i = 0 To aux2
            ws.Cells(Krang.Row + i, ws.Range("MEL_LST[[#Headers],[Control]]").Column).Value = "Deleted"
        Next
    
    End If
        
End If
        
        Application.EnableEvents = True

wb.Save

Exit Function

error5: MsgBox ("error5: Procedure - change in control column has failed")
        Application.EnableEvents = True
        Set wb = Nothing
        Set ws = Nothing
        Call start
      
Deleted: MsgBox ("MEL owner has deleted row. It's all good ;oD")
        Application.EnableEvents = True
        wb.Save
        Call start
   
End Function

Function change_revision(ByVal Trow As Range, ByVal Krang As Range, net As String)

On Error GoTo error6
       
Dim validation As Boolean
Dim cell As Range
       
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        If WorksheetFunction.XLookup(Trow.Value, wb.Sheets("VARIANCES").Range("V_REV[REVA]"), wb.Sheets("VARIANCES").Range("V_REV[REVN]")) _
        > WorksheetFunction.XLookup(wb.Sheets("BUP").Range("REVBUP").Value, wb.Sheets("VARIANCES").Range("V_REV[REVA]"), wb.Sheets("VARIANCES").Range("V_REV[REVN]")) Then
            
            'add code to double check that the MEL will be rev-up
            validation = MsgBox("Are you sure you want to rev-up the MEL?", vbOKCancel, "MEL REV-UP") = vbOK
                    
            If validation = True Then
            
                For Each cell In ws.Cells.Range("MEL_LST[CONTROL]")
                    If cell.Value <> "Deleted" Then
                        cell.ClearContents
                    End If
                
                Next
                ws.Range("MEL_LST[#data]").Interior.ColorIndex = 0
                wb.Sheets("BUP").Range("REVBUP").Value = Trow.Value
            Else
            
                     MsgBox ("The MEL hasn't been reved-up")
                    Trow.Value = wb.Sheets("BUP").Range("REVBUP").Value
            
            End If
        
        ElseIf WorksheetFunction.XLookup(Trow.Value, wb.Sheets("VARIANCES").Range("V_REV[REVA]"), wb.Sheets("VARIANCES").Range("V_REV[REVN]")) _
        < WorksheetFunction.XLookup(wb.Sheets("BUP").Range("REVBUP").Value, wb.Sheets("VARIANCES").Range("V_REV[REVA]"), wb.Sheets("VARIANCES").Range("V_REV[REVN]")) And net = "Abel" Then
        
            validation = MsgBox("Are you sure you want to rev-down the MEL?. This is a very unnormal procedure", vbOKCancel, "MEL REV-UP") = vbOK
        
            If validation = True Then
                    MsgBox ("The change control format won't be modified")
                    wb.Sheets("BUP").Range("REVBUP").Value = Trow.Value
                        
            Else
                    MsgBox ("The MEL hasn't been reved-down")
                    Trow.Value = wb.Sheets("BUP").Range("REVBUP").Value
            
            End If
        
        Else
        
        MsgBox ("The MEL hasn't been reved-up")
        Trow.Value = wb.Sheets("BUP").Range("REVBUP").Value
        
        End If
                
        Call cellBlock
                
                
        Application.EnableEvents = True
        Application.ScreenUpdating = True

wb.Save

Exit Function

error6: MsgBox ("error6: Procedure - change in revision has failed")
        Call cellBlock
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Set wb = Nothing
        Set ws = Nothing
        Call start

End Function

Function change_mode(ByVal Trow As Range, ByVal Krang As Range, user As String)


On Error GoTo error7

Application.ScreenUpdating = False
Application.EnableEvents = False

If access = 1 Then
            
    Select Case ws.Range("A3").Value
    
        Case "MAINTENANCE"
            'unprotect sheet
           ws.Unprotect Password:=pswd
            'hide all buttons
            Me.Buttons("Button 38").Enabled = False
            Me.Buttons("Button 39").Enabled = False
            ws.Range("B1").Value = "MEL IS UNDERGOING MAINTENANCE, CONTACT @ABEL FOR MORE INFORMATION"
            
           Application.Interactive = False
           'AQUIII add code to avoid anyone from making any change
        
        Case ""
            'unhide all buttons
            Me.Buttons("Button 38").Enabled = True
            Me.Buttons("Button 39").Enabled = True
            ws.Range("B1").Value = ""
            ws.Protect Password:=pswd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Userinterfaceonly:=True, AllowFormattingColumns:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
            
            Application.Interactive = True
        Case "PAUSE"
            'unprotect sheet
            ws.Unprotect Password:=pswd
            
            'hide all buttons
            Me.Buttons("Button 38").Enabled = False
            Me.Buttons("Button 39").Enabled = False
            ws.Range("B1").Value = "MEL IS TEMPORARY CLOSED, CONTACT THE PROCESS DEPARTMENT FOR MORE INFORMATION"
            
            
            Application.Interactive = True
            'AQUIII add code to avoid anyone excep / process from making any change
            
    End Select


ElseIf access = 2 Then

    Select Case ws.Range("A3").Value
    
        Case ""
            'unhide all buttons
            Me.Buttons("Button 38").Enabled = True
            Me.Buttons("Button 39").Enabled = True
            ws.Range("B1").Value = ""
            
        Case "PAUSE"
            'hide all buttons
            Me.Buttons("Button 38").Enabled = False
            Me.Buttons("Button 39").Enabled = False
            ws.Range("B1").Value = "MEL IS TEMPORARY CLOSED, CONTACT THE PROCESS DEPARTMENT FOR MORE INFORMATION"
            
             'AQUIII add code to avoid anyone excep / process from making any change
        Case Else
            
            Sheets("MEL").Range("MODE").Value = ""
            MsgBox ("You don't have permissions to change the mode, please contact abel@halyard.ca")
            Me.Buttons("Button 38").Enabled = True
            Me.Buttons("Button 39").Enabled = True
            ws.Range("B1").Value = ""
            
    End Select


End If

Application.ScreenUpdating = True
Application.EnableEvents = True

wb.Save

Exit Function

error7: MsgBox ("error7: Procedure - change in mode has failed")
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Interactive = True
        Call start

End Function

Function type_selection(ByVal Trow As Range, ByVal Krang As Range, net As String)
    
On Error GoTo error8

Dim validation As Boolean
Dim aux As Integer


Application.EnableEvents = False

If ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[WBS]]").Column).Value <> "" Then

Dim network As Object
Set network = CreateObject("WScript.Network")
  
Dim tbl As ListObject
Set tbl = ws.ListObjects("MEL_LST")


'AQUI codigo para controlar el cambio de tipo.
'Reflexionar sobre los cambios que se pueden producir y afectar a valores que no estaràn controlados por formulas.
'es posible que se requira un reseteo de las formulas

'AQUI Cambio de typo solo serà posible en modo START.
'Modificaciones adicionales solo podràn realizarse eliminando elementos y aÐadiendolos al final.

        
        
        
        
If ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[MOTOR]]").Column).Value <> "Y" And Krang.Columns.Count = 1 Then

    Select Case Krang.Value

    Case "AG"
    
        validation = MsgBox("Do you want to add a tank with this agitator?", vbOKCancel, "ADDING AN ASSOCIATED TANK") = vbOK
    
        If validation = True Then

            
            aux = ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[MOTOR QUANTITY]]").Column).Value
    
            If Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[MOTOR QUANTITY]]").Column).Value = 1 And Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[EMERGENCY LOAD]]").Column).Value = "---" Then
    
                tbl.ListRows.Add Position:=Krang.Row - 5 + aux
    
            ElseIf ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[EMERGENCY LOAD]]").Column).Value = "-" Then
    
                tbl.ListRows.Add Position:=Krang.Row - 5
                aux = 0
    
            Else
    
                tbl.ListRows.Add Position:=Krang.Row - 5 + aux
    
            End If
            
            ' code to fill in the new line in the table iwith the right information
            ws.Cells(Krang.Row + 1 + aux, ws.Range("MEL_LST[[#Headers],[DATE]]").Column).Value = Format(Date, "yyyy/mm/dd")
            ws.Cells(Krang.Row + 1 + aux, ws.Range("MEL_LST[[#Headers],[WBS]]").Column).Value = _
            ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[WBS]]").Column).Value
            ws.Cells(Krang.Row + 1 + aux, ws.Range("MEL_LST[[#Headers],[TYPE]]").Column).Value = "TK"
            ws.Cells(Krang.Row + 1 + aux, ws.Range("MEL_LST[[#Headers],[CONTROL]]").Column).Value = network.UserName
            ws.Cells(Krang.Row + 1 + aux, ws.Range("MEL_LST[[#Headers],[NUMBER]]").Column).Value = Range("MEL_ROWS").Value
                    
            
            If ws.Range("Version").Value = "START" Then
    
                ws.Cells(Krang.Row + 1 + aux, ws.Range("MEL_LST[[#Headers],[REV]]").Column).Value = "A"
    
            Else
    
                ws.Cells(Krang.Row + 1 + aux, ws.Range("MEL_LST[[#Headers],[REV]]").Column).Value = ws.Range("VERSION").Value
    
            End If
            
        Else
            MsgBox ("No additional items have been added")
        
        End If
    
    
    Case "TH"
    
        validation = MsgBox("Do you want to add a thickener mechanism for this thickener?", vbOKCancel, "ADDING AN ASSOCIATED TANK") = vbOK
    
        If validation = True Then
            'code to add new line in the table in the right position
    
            If Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[MOTOR QUANTITY]]").Column).Value = 1 And Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[EMERGENCY LOAD]]").Column).Value = "---" Then
    
                tbl.ListRows.Add Position:=Krang.Row - 5 + aux
    
            ElseIf ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[EMERGENCY LOAD]]").Column).Value = "-" Then
    
                tbl.ListRows.Add Position:=Krang.Row - 5
                aux = 0
    
            Else
    
                tbl.ListRows.Add Position:=Krang.Row - 5
    
            End If
                      
   
            'code to fill in the new line in the table iwith the right information
            ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[DATE]]").Column).Value = Format(Date, "yyyy/mm/dd")
            ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[WBS]]").Column).Value = _
            ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[WBS]]").Column).Value
            ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[TYPE]]").Column).Value = "TM"
            ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[CONTROL]]").Column).Value = network.UserName
            ws.Cells(Krang.Row + 1 + aux, ws.Range("MEL_LST[[#Headers],[NUMBER]]").Column).Value = Range("MEL_ROWS").Value
            
             If ws.Range("Version").Value = "START" Then
    
                ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[REV]]").Column).Value = "A"
    
            Else
    
                ws.Cells(Krang.Row + 1, ws.Range("MEL_LST[[#Headers],[REV]]").Column).Value = ws.Range("VERSION").Value
    
            End If
            
            
        Else
            MsgBox ("No additional items have been added")
        
        End If
    
    
    End Select

End If

Else

MsgBox ("Please, select first a wbs code")
ws.Cells(Krang.Row, ws.Range("MEL_LST[[#Headers],[TYPE]]").Column).Value = ""

End If
        
        Application.EnableEvents = True
        Set tbl = Nothing
        Set network = Nothing

Exit Function

error8: MsgBox ("error8: Procedure - add associated equipment in table has failed")
        Application.EnableEvents = True
        Set tbl = Nothing
        Set network = Nothing
        Call start

End Function

Sub change_description(ByVal Trow As Range, ByVal Krang As Range, net As String)

Application.EnableEvents = False

On Error GoTo error9


Application.EnableEvents = True

Exit Sub

error9: MsgBox ("error9: Procedure - sorted item in table has failed")
        Application.EnableEvents = True
        'Set tbl = Nothing
        Application.EnableEvents = True
        Call start

End Sub

Sub chage_stepstatus(ByVal Trow As Range, ByVal Krang As Range, net As String)

Application.EnableEvents = False

On Error GoTo error40

LockTableRowsByStatus

Application.EnableEvents = True

Exit Sub

error40: MsgBox ("error40: Procedure - change STEP STATUS has failed")
        Application.EnableEvents = True
        'Set tbl = Nothing
        Application.EnableEvents = True
        Call start

End Sub


