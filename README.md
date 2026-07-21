# my-memo-text
Private Sub UploadToSharePoint(filePath As String, fileName As String)


    Dim http As Object
    Dim stream As Object
    Dim url As String


    url = _
    "https://company.sharepoint.com/sites/MailArchive/Documents/MailBackup/" _
    & fileName



    Set stream = CreateObject("ADODB.Stream")


    stream.Type = 1

    stream.Open

    stream.LoadFromFile filePath



    Set http = CreateObject("MSXML2.XMLHTTP")


    http.Open "PUT", url, False


    http.setRequestHeader _
        "Content-Type", _
        "application/octet-stream"



    http.Send stream.Read



    stream.Close



    If http.Status <> 200 _
       And http.Status <> 201 Then


        MsgBox "SharePoint保存失敗：" _
        & http.Status _
        & vbCrLf _
        & http.responseText


    End If


End Sub
