Global dim VBApswd as string
Public Const UPDATE_MESSAGE As String = "New version installed. See release notes for details."

Sub updates()

On Error GoTo updatefail

    Dim moduleURL As String
    Dim objectURL As String
    Dim tempFolder As String
    Dim vbComp As Object
    Dim tmpFile As String
    Dim fileURL As String
    Dim compName As String
    Dim changeMade As Boolean

    moduleURL = "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MODULES/"
    objectURL = "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MICROSOFT_EXCEL_OBJECTS/"
    tempFolder = Environ("TEMP") & "\"

    changeMade = False

    UnlockVBProject VBApswd

    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        fileURL = ""
        tmpFile = ""

        Select Case vbComp.Type
            Case 1
                If vbComp.Name <> "m_update" Then
                    compName = vbComp.Name
                    fileURL = moduleURL & compName & ".bas"
                    tmpFile = tempFolder & compName & ".bas"
                    If DownloadFile(fileURL, tmpFile) Then
                        If ComponentChanged(vbComp, tmpFile, tempFolder) Then
                            changeMade = True
                            ThisWorkbook.VBProject.VBComponents.Remove vbComp
                            Set vbComp = ThisWorkbook.VBProject.VBComponents.Import(tmpFile)
                            vbComp.Name = compName
                        End If
                    Else
                        GoTo updatefail
                    End If
                End If
            Case 100
                fileURL = objectURL & vbComp.Name & ".cls"
                tmpFile = tempFolder & vbComp.Name & ".cls"
                If DownloadFile(fileURL, tmpFile) Then
                    If ComponentChanged(vbComp, tmpFile, tempFolder) Then
                        changeMade = True
                        If vbComp.CodeModule.CountOfLines > 0 Then
                            vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
                        End If
                        vbComp.CodeModule.AddFromFile tmpFile
                    End If
                Else
                    GoTo updatefail
                End If
        End Select

    Next vbComp

    If changeMade Then
        MsgBox UPDATE_MESSAGE, vbInformation
    End If

    Exit Sub

updatefail:
    MsgBox "Unable to retrieve latest code. Please contact abel@halyard.ca", vbCritical
    Application.DisplayAlerts = False
		call start
    ThisWorkbook.Close SaveChanges:=False
    End

End Sub

Private Sub UnlockVBProject(ByVal password As String)
    Dim vbp As Object
    Set vbp = ThisWorkbook.VBProject

    If vbp.Protection <> 0 Then
        Application.VBE.MainWindow.Visible = True
        vbp.VBE.CommandBars("Menu Bar").Controls("Tools").Controls("VBAProject Properties...").Execute
        Application.SendKeys password & "{ENTER}", True
        DoEvents
        Application.VBE.MainWindow.Visible = False
    End If
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
		call start

End Function

Private Function ComponentChanged(ByVal vbComp As Object, ByVal downloadedFile As String, ByVal tempFolder As String) As Boolean
    Dim fso As Object
    Dim localFile As String
    localFile = tempFolder & "local_tmp"
    vbComp.Export localFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    ComponentChanged = (ReadFile(localFile) <> ReadFile(downloadedFile))
    fso.DeleteFile localFile, True
End Function

Private Function ReadFile(ByVal path As String) As String
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(path, 1)
    ReadFile = ts.ReadAll
    ts.Close
End Function

