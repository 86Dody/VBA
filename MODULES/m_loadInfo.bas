Attribute VB_Name = "m_loadInfo"
Option Explicit

Sub loaddata()


On Error GoTo error10

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim fileD As FileDialog
    Dim f_file As Variant
    Dim Filepath() As String
    Dim Filepath_b() As Boolean
    Dim f_ext() As String
    Dim FSO As Object
    Dim nummern As Integer
    Dim err_ext As Integer
    
    Dim wbM As Workbook
    Dim wb1 As Workbook
    
    Dim wsM As Worksheet
    Dim ws1 As Worksheet
    Dim cell As Range
    
    Dim nm() As String
    
    Dim i As Integer
    Dim ii As Integer
    
    Set wbM = ActiveWorkbook
    Set wsM = wbM.Sheets("MEL")
    
    ii = 1
    err_ext = 0
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Set fileD = Application.FileDialog(msoFileDialogFilePicker)
    fileD.AllowMultiSelect = True
    
    With fileD
        .Title = "Select the Datasheet '.xls' file(s) you wish to load information from."
        .Show
        
        If .SelectedItems.Count <> 0 Then
            ReDim Preserve Filepath(.SelectedItems.Count)
            ReDim Preserve Filepath_b(.SelectedItems.Count)
            ReDim Preserve f_ext(.SelectedItems.Count)
            ReDim Preserve nm(.SelectedItems.Count)
            For Each f_file In fileD.SelectedItems
                Filepath(ii) = fileD.SelectedItems(ii)
                ii = ii + 1
            Next
            nummern = fileD.SelectedItems.Count
        Else
            'Code to handle event that nothing is selected
            'e.g.
            Exit Sub
        End If
    End With

 'check if the file(s) has the addecuate extension .xls
 
    For i = 1 To nummern
    
    f_ext(i) = FSO.GetExtensionName(Filepath(i))
    nm(i) = FSO.getfilename(Filepath(i))
        
            If f_ext(i) <> "xlsx" Then
                
                Filepath_b(i) = False
                err_ext = err_ext + 1
            Else
                
                Filepath_b(i) = True
                'Open the file
                Set wb1 = Workbooks.Open(Filepath(i))
                Set ws1 = wb1.Sheets(1)
                
                'look for the TAG
                'DUTY / SIZE
                'MODEL
                'WEIGHT (Kg)
                'POWER [HP]
                'POWER [kW]
                'MOTOR QUANTITY
                'VOLTS [V]
                'FREQUENCY [Hz]
                'PHASE
                
                'Sheets("MEL").Range ("MEL_LST[[EQUIPMENT DESCRIPTION]:[PHASE]]")
                              
                'find the row in the table
                Dim r As Range
                Dim rr As Range
                Dim rrrr As Integer
                
                Set r = ws1.Range("TAG")
                Set rr = wsM.Range("MEL_LST[TAG]")
                
                For Each cell In rr
                
                    If cell.Value = r.Value Then
                    
                        rrrr = cell.Row
                    
                    End If
                    
                
                Next
                'write the parameters from the DS to the MEL
                
                Application.EnableEvents = True
                wsM.Cells(rrrr, wsM.Range("MEL_LST[[#Headers],[DUTY / SIZE]]").Column).Value = ws1.Range("DUTY___SIZE")
                wsM.Cells(rrrr, Range("MEL_LST[[#Headers],[MODEL]]").Column).Value = ws1.Range("MODEL").Value
                wsM.Cells(rrrr, Range("MEL_LST[[#Headers],[WEIGHT (Kg)]]").Column).Value = ws1.Range("WEIGHT__Kg").Value
                
                'AQUIIIII tengo que codificar donde colocar la informaci—n de potencia. Esta deber’a ir en los motores solamente
                
                'wsM.Cells(rrrr, Range("MEL_LST[[#Headers],[POWER (HP)]]").Column).Value = ws1.Range("POWER__HP").Value
                'wsM.Cells(rrrr, Range("MEL_LST[[#Headers],[POWER (kW)]]").Column).Value = ws1.Range("POWER__kW").Value
                'wsM.Cells(rrrr, Range("MEL_LST[[#Headers],[MOTOR QUANTITY]]").Column).Value = ws1.Range("MOTOR_QUANTITY").Value
                
                wsM.Cells(rrrr, Range("MEL_LST[[#Headers],[VOLTS (V)]]").Column).Value = ws1.Range("VOLTS__V").Value
                'wsM.Cells(rrrr, Range("MEL_LST[[#Headers],[FREQUENCY [Hz]]]").Column).Value = ws1.Range("FREQUENCY__Hz").Value
                'wsM.Cells(rrrr, Range("MEL_LST[[#Headers],[PHASE]]").Column).Value = ws1.Range("PHASE").Value
                Application.EnableEvents = False
                
                'close the datasheet
                wb1.Close
                
                'process completed
                MsgBox ("The transfer of information has been completed")
                
            End If

    Next

    If err_ext <> 0 Then
    
        MsgBox (err_ext & " files don't have the expected '.xlsx' extension" & vbNewLine & "For those items," & _
        "no information has been transferred")
    
    End If

Set FSO = Nothing
Set ws1 = Nothing
Set wb1 = Nothing

Application.ScreenUpdating = True
'Application.EnableEvents = True

Exit Sub

error10:

MsgBox ("error10: Procedure - load data from DS has failed")
Call start
Set FSO = Nothing
Set ws1 = Nothing
Set wb1 = Nothing

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
