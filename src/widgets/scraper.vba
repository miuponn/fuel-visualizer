Sub DownloadFile()

Dim myURL As String
myURL = "https://YourWebSite.com/?your_query_parameters"

Dim WinHttpReq As Object
Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", myURL, False, "username", "password"
WinHttpReq.send

If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.responseBody
    oStream.SaveToFile "C:\file.csv", 2 ' 1 = no overwrite, 2 = overwrite
    oStream.Close
End If
