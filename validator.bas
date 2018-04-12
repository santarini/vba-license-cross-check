Sub DateTest()

'enable MSFT scripting runtime

Dim x, z As Long
Dim Reference As Variant

Dim Json As Object

'create a dicitonary
Dim Dict As New Dictionary
Dict.CompareMode = CompareMethod.TextCompare

Dict("A") = Environ("computername")
Dict("B") = Environ("username")


Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
MyRequest.Open "GET", "http://localhost"
MyRequest.Send

Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)

Reference = Json("testCompany")(Dict.Item("A"))(Dict.Item("B"))("finSoft")("licenses")("references")

z = DateValue(Date)

MsgBox z

If InStr(1, Reference, z) = 0 Then
    MsgBox "Your License is Expired."
Else
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
    If IsEmpty(sendResult) = True Then
        MsgBox "No Internet Connection"
    End If
End Sub
