Attribute VB_Name = "ConsumerBasic"
Option Explicit

' Example for a workbook that references the verified VBA-HTTP release artifact.
Public Sub GetLoopbackStatus()
    Dim client As HttpClient
    Dim response As HttpResponse

    Set client = VBAHttp.CreateClient()
    Set response = client.GetResponse("http://127.0.0.1:8080/status/204")
    response.RaiseForStatus
    Debug.Print "status=" & CStr(response.StatusCode)
End Sub

Public Sub GetJsonWithHeaders()
    Dim client As HttpClient
    Dim request As HttpRequest
    Dim response As HttpResponse

    Set client = VBAHttp.CreateClient()
    Set request = VBAHttp.CreateRequest()
    request.Method = "GET"
    request.Url = "https://example.test/api"
    request.Headers.SetValue "Accept", "application/json"
    Set response = client.Execute(request)
    response.RaiseForStatus
    Debug.Print response.Text
End Sub
