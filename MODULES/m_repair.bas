
Sub repair()
'HOLA
Application.EnableEvents = False

'path of the reference file
Dim pthref As String
pthref = "C:\Users\Abel\OneDrive - Halyard Inc\Desktop\H24376-0000-M-MEL-001_REVIEW.xlsx"

Dim ws As Worksheet
Dim wb As Workbook

Set wb = Workbooks("M25058-0000-M-MEL-001 - MECHANICAL EQUIPMENT LIST_test.xlsm")
Set ws = wb.Sheets("MEL")

Dim wsref As Worksheet
Dim wbref As Workbook

Set wbref = Workbooks.Open(pthref)
Set wsref = wbref.Sheets("Sheet1")

Dim tb As ListObject
Dim tbref As ListObject

Set tb = ws.ListObjects("MEL_LST")
Set tbref = wsref.ListObjects("Table3")

Dim tb_row As ListRow
Dim tbref_row As ListRow

Dim tb_col As ListColumn
Dim tbref_col As ListColumn

For Each tbref_row In tbref.ListRows
    
    For Each tbref_col In tbref.ListColumns
        
        'check if formula
        'If WorksheetFunction.IsFormula(ws.Cells(tbref_row.Index + 6, tbref_col.Index + 1)) = False Then
        'If tbref_col.Index = 32 Then
        
            If wsref.Cells(tbref_row.Index + 6, tbref_col.Index + 1).Value <> "-" And Cells(tbref_row.Index + 6, tbref_col.Index + 1).Value <> "---" And Cells(tbref_row.Index + 6, tbref_col.Index + 1).Value <> "" Then
            
                ws.Cells(tbref_row.Index + 6, tbref_col.Index + 1).Value = wsref.Cells(tbref_row.Index + 6, tbref_col.Index + 1).Value
                
            End If
        
       'End If
    
    Next

Next

Application.EnableEvents = True

End Sub

