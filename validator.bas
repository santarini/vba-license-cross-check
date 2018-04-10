Sub DateTest()

'enable MSFT scripting runtime

Dim x As Long
Dim y, z As Date
Dim reference As Variant

Dim Json As Object

'create a dicitonary
'Dim Dict As New Dictionary
'Dict.CompareMode = CompareMethod.TextCompare

Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
MyRequest.Open "GET", "http://localhost"
MyRequest.Send

Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)

reference = Json("testCompany")("finSoft")("licenses")("references")

MsgBox reference

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
