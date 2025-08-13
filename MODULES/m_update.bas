Attribute VB_Name = "m_update"

Sub updates()
On Error GoTo updatefail

    Dim moduleURL As String
    Dim objectURL As String
    Dim tempFolder As String
    Dim vbComp As Object
    Dim tmpFile As String
    Dim fileURL As String

    moduleURL = "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MODULES/"
    objectURL = "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MICROSOFT_EXCEL_OBJECTS/"
    tempFolder = Environ("TEMP") & "\"

    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        fileURL = ""
        tmpFile = ""

        Select Case vbComp.Type
            Case 1
                If vbComp.Name <> "m_update" Then
                    fileURL = moduleURL & vbComp.Name & ".bas"
                    tmpFile = tempFolder & vbComp.Name & ".bas"
                    If DownloadFile(fileURL, tmpFile) Then
                        ThisWorkbook.VBProject.VBComponents.Remove vbComp
                        ThisWorkbook.VBProject.VBComponents.Import tmpFile
                    Else
                        GoTo updatefail
                    End If
                End If
            Case 100
                fileURL = objectURL & vbComp.Name & ".cls"
                tmpFile = tempFolder & vbComp.Name & ".cls"
                If DownloadFile(fileURL, tmpFile) Then
                    vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
                    vbComp.CodeModule.AddFromFile tmpFile
                Else
                    GoTo updatefail
                End If
        End Select

    Next vbComp

    Exit Sub

updatefail:
    MsgBox "Unable to retrieve latest code. Please contact abel@halyard.ca", vbCritical
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=False
    End

End Sub

Private Function DownloadFile(ByVal url As String, ByVal dest As String) As Boolean
On Error GoTo errHandler

    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send

    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write http.responseBody
        stream.SaveToFile dest, 2
        stream.Close
        Set stream = Nothing
        Set http = Nothing
        DownloadFile = True
    Else
        DownloadFile = False
    End If

    Exit Function

errHandler:
    DownloadFile = False

End Function

