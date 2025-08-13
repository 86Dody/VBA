Attribute VB_Name = "m_update"
Option Explicit

Private Const MODULE_URL As String = _
    "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MODULES/"
Private Const OBJECT_URL As String = _
    "https://halyardinc-my.sharepoint.com/:f:/r/personal/abel_halyard_ca/Documents/Documents/Abel/Programing/GitHub/VBA/MICROSOFT_EXCEL_OBJECTS/"

Sub updates()
    On Error GoTo updatefail

    Dim tempFolder As String
    Dim vbComp As Object
    Dim tmpFile As String
    Dim fileURL As String

    tempFolder = Environ("TEMP") & "\\"

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        fileURL = ""
        tmpFile = ""

        Select Case vbComp.Type
            Case 1
                If vbComp.Name <> "m_update" Then
                    fileURL = MODULE_URL & vbComp.Name & ".bas"
                    tmpFile = tempFolder & vbComp.Name & ".bas"
                    If DownloadFile(fileURL, tmpFile) Then
                        If Not CodeMatchesFile(vbComp, tmpFile) Then
                            ThisWorkbook.VBProject.VBComponents.Remove vbComp
                            ThisWorkbook.VBProject.VBComponents.Import tmpFile
                        End If
                        On Error Resume Next
                        Kill tmpFile
                        On Error GoTo updatefail
                    Else
                        GoTo updatefail
                    End If
                End If
            Case 100
                fileURL = OBJECT_URL & vbComp.Name & ".cls"
                tmpFile = tempFolder & vbComp.Name & ".cls"
                If DownloadFile(fileURL, tmpFile) Then
                    If Not CodeMatchesFile(vbComp, tmpFile) Then
                        vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
                        vbComp.CodeModule.AddFromFile tmpFile
                    End If
                    On Error Resume Next
                    Kill tmpFile
                    On Error GoTo updatefail
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

Private Function CodeMatchesFile(vbComp As Object, ByVal filePath As String) As Boolean
    Dim fileCode As String
    Dim currentCode As String

    fileCode = GetCodeFromFile(filePath)
    currentCode = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)

    CodeMatchesFile = (StrComp(currentCode, fileCode, vbBinaryCompare) = 0)
End Function

Private Function GetCodeFromFile(ByVal filePath As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim content As String
    Dim startPos As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)
    content = ts.ReadAll
    ts.Close

    startPos = InStr(1, content, "Attribute VB_Name")
    If startPos > 0 Then
        content = Mid$(content, startPos)
    End If

    GetCodeFromFile = content
End Function
