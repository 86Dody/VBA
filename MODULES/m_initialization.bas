

'define all public variables
Global pswd As String
Global access As Integer
Global secure As Range


Sub start()

On Error GoTo error1
Application.ScreenUpdating = False
Application.EnableEvents = False

Dim network As Object
Set network = CreateObject("WScript.Network")

pswd = "ConPdeProcess"

Dim wb As Workbook
Dim ws As Worksheet
Dim aer As AllowEditRange


'erase existing permissions
For Each ws In Worksheets
   ws.Unprotect Password:=pswd
        For Each aer In ws.Protection.AllowEditRanges
            aer.Delete
        Next aer
Next ws

Set wb = Workbooks("M25058-2000-M-MEL-003.xlsm")
Set ws = wb.Sheets("MEL")

'Cells(1, 1).Value = network.UserName

'define acess levels

Select Case network.UserName
    
    'item to give temporary access as owner
    Case ws.Range("ADM").Value
    
    access = 3
    
    Case "Abel"
    
    access = 1
       
    Case "ivana"
    
    access = 2
    
    Case "Valan"
    
    access = 2
    
    Case "Michael"
    
    access = 2
       
    Case Else
    
    access = 3
    
End Select

If access <= 2 Then
    
    If ws.Protection.AllowEditRanges.Count < 7 Then
    'giving permissions for control column
    ws.Protection.AllowEditRanges.Add Title:="RControl", Range:=Range("MEL_LST[CONTROL]")
    'giving permissions for MEL main data (Version, data, owner, project name and number...)
    ws.Protection.AllowEditRanges.Add Title:="Vcontrol", Range:=Range("VERSION")
    'giving permissions for MEL main data (Version, data, owner, project name and number...)
    ws.Protection.AllowEditRanges.Add Title:="Vmode", Range:=Range("MODE")
    'giving permissions to modify the temporary admin
    ws.Protection.AllowEditRanges.Add Title:="Vadm", Range:=Range("ADM")
    'giving permissions to modify the temporary STEP2A
    ws.Protection.AllowEditRanges.Add Title:="STEP2A", Range:=Range("STEP2A")
    'giving permissions to modify the temporary STEP2B
    ws.Protection.AllowEditRanges.Add Title:="STEP2B", Range:=Range("STEP2B")
    'giving permissions to modify the temporary STEP3
    ws.Protection.AllowEditRanges.Add Title:="STEP3", Range:=Range("STEP3")
    End If

Else



End If

    With Application
        .Iteration = False
        .Calculation = xlCalculationAutomatic
    End With
   
'be sure that the order is correct.
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
   
'applied the adequate cell locking configuration for each user
Call LockTableRowsByStatus
   
ws.Protect Password:=pswd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Userinterfaceonly:=True, AllowFormattingColumns:=True, AllowInsertingRows:=True, AllowDeletingRows:=True

Call cellBlock

Application.ScreenUpdating = True
Application.EnableEvents = True

Set ws = Nothing
Set wb = Nothing
Set aer = Nothing
Set network = Nothing

Exit Sub

error1:      MsgBox ("error1: Procedure - initialization has failed")
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Set ws = Nothing
            Set aer = Nothing
            Set network = Nothing
            
End Sub

Public Sub LockTableRowsByStatus()

On Error GoTo error102

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("MEL_LST") ' Modify if needed

    Dim frozenRange As Range
    Dim certifiedRange As Range
    Dim rowRange As Range
    Dim stepName As String
    Dim stepStatus As String
    Dim statusColumnIndex As Long

    'store standard lock cell
    SelectLockedCells

    ' Find the STATUS column
    On Error Resume Next
    statusColumnIndex = tbl.ListColumns("STATUS").Index
    On Error GoTo error102
    If statusColumnIndex = 0 Then
        MsgBox "STATUS column not found.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    
   ws.Unprotect Password:=pswd

    ' First unlock all rows
    tbl.DataBodyRange.Locked = False

    ' Loop through each row
    For Each rowRange In tbl.DataBodyRange.Rows
        stepName = Trim(UCase(rowRange.Cells(1, statusColumnIndex).Value))
        
        ' Get status for that step
        Select Case stepName
            Case "STEP2A"
                stepStatus = Trim(UCase(Range("STEP2A").Value))
            Case "STEP2B"
                stepStatus = Trim(UCase(Range("STEP2B").Value))
            Case "STEP3"
                stepStatus = Trim(UCase(Range("STEP3").Value))
            Case Else
                stepStatus = ""
        End Select
        
        ' Apply logic based on status and access level
        Select Case stepStatus
            Case "FROZEN"
                If access <> 1 And access <> 2 Then
                    If frozenRange Is Nothing Then
                        Set frozenRange = rowRange
                    Else
                        Set frozenRange = Union(frozenRange, rowRange)
                    End If
                End If
            Case "CERTIFIED"
                If access <> 1 Then
                    If certifiedRange Is Nothing Then
                        Set certifiedRange = rowRange
                    Else
                        Set certifiedRange = Union(certifiedRange, rowRange)
                    End If
                End If
            Case "NORMAL"
                ' Do nothing
            Case Else
                ' Unknown status — optionally handle
        End Select
        
        'color formatting
        
        With rowRange.Cells(1, statusColumnIndex)
            Select Case stepStatus
                Case "FROZEN"
                    .Interior.Color = RGB(189, 215, 238) ' Clear blue
                Case "CERTIFIED"
                    .Interior.Color = RGB(198, 239, 206) ' Clear green
                Case "NORMAL"
                    .Interior.ColorIndex = xlNone
                Case Else
                    .Interior.Color = RGB(255, 199, 206) ' Optional: highlight as warning
            End Select
        End With

        
    Next rowRange

    ' Apply locking
    secure.Locked = True
    If Not frozenRange Is Nothing Then frozenRange.Locked = True
    If Not certifiedRange Is Nothing Then certifiedRange.Locked = True

    ws.Protect Password:=pswd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Userinterfaceonly:=True, AllowFormattingColumns:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
    Application.ScreenUpdating = True

    
Exit Sub

Set ws = Nothing
Set tbl = Nothing

error102:
        Set ws = Nothing
        Set tbl = Nothing
        MsgBox ("error102: Procedure - the cell blocking has failed")
        Call start
End Sub

Public Sub SelectLockedCells()

On Error GoTo error101

'storing the standard cell lock configuration
    
Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("MEL_LST")

    Dim colNames As Variant
    colNames = Array("CONTROL", "REV", "DATE", "CERTIFIED", "NUMBER", "TYPE", "TAG", _
                     "SUPPLY PKG DESCRIPTION", "SUPPLIER", "MOTOR QUANTITY", _
                     "POWER [AVERAGE]", "KVA", "INDEX", "INDEX_TAG", "MOTOR", "MOTORIND")

    Dim targetRange As Range
    Dim col As ListColumn
    Dim i As Long

    ' Loop through desired column names
    For i = LBound(colNames) To UBound(colNames)
        On Error Resume Next
        Set col = tbl.ListColumns(colNames(i))
        On Error GoTo 0

        If Not col Is Nothing Then
            If targetRange Is Nothing Then
                Set targetRange = col.DataBodyRange
            Else
                Set targetRange = Union(targetRange, col.DataBodyRange)
            End If
        End If
    Next i

    If Not targetRange Is Nothing Then
        targetRange.Select
    Else
        MsgBox "None of the specified columns were found in table MEL_LST.", vbExclamation
    End If
    
    Set secure = Selection
    Range("MEL_LST[[#Headers],[CONTROL]]").Select
    
Exit Sub

error101:
        MsgBox ("error101: Procedure - Defining standard cell locking configuration has failed, close the file and contact Abel")
        Call start

End Sub


