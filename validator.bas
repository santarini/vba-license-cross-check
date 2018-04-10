Sub DateTest()

Dim x As Long
Dim y, z As Date


x = 43199
y = Format(x, "Short Date")
z = Date

If y = z Then
    MsgBox "Valid License"
End If

End Sub

Sub CheckInternetConnection()
    Dim sendResult As String
    Dim objHTTP As Object
    Dim URL As String
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    URL = "http://www.google.com"
    objHTTP.Open "GET", URL, False
    objHTTP.Send
    sendResult = objHTTP.ResponseText
    MsgBox sendResult
End Sub
