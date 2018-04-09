Sub DateTest()

Dim x As Long
Dim y As Date


x = 43191

y = Format(x, "Short Date")

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
