Sub DateTest()

Dim x As Long
Dim y As Date


x = 43191

y = Format(x, "Short Date")

End Sub



Sub checkInternet()

On Error Resume Next
Dim request As MSXML2.XMLHTTP60
request.Open "http://www.google.com"
request.Send
MsgBox request.Status

End Sub
