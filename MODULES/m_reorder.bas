
Sub reorder()

'this function is not being used

Application.EnableEvents = False

On Error GoTo error25

Dim ws As Worksheet
Set ws = Sheets("MEL")

If WorksheetFunction.CountIf(Range("MEL_LST[EQUIPMENT DESCRIPTION]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TAG]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[WBS]"), "") = 0 And WorksheetFunction.CountIf(Range("MEL_LST[TYPE]"), "") = 0 Then

ws.Unprotect Password:=pswd

Dim aux As Integer

ws.ListObjects("MEL_LST").Sort.SortFields.Clear
'Removing re-order to allow for a different tag numbering

ws.ListObjects("MEL_LST").Sort.SortFields.Add2 _
        Key:=Range("MEL_LST[NUMBER]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortTextAsNumbers

With ws.ListObjects("MEL_LST").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With

Else

MsgBox ("Before ordering the equipment proceed to complete the missing information (WBS, Type, Description)")

End If



Application.EnableEvents = True

Exit Sub

error25:     MsgBox ("error25: Procedure - sorted item in table has failed")
        Application.EnableEvents = True
        Set tbl = Nothing
        Set ws = Nothing
        Application.EnableEvents = True
        Call start

End Sub


