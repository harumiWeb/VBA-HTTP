Attribute VB_Name = "WinHttpComTransportTests"
Option Explicit

'@Tag("integration")
Public Sub Test_ComTransport_GetMapsQueryHeadersAndResponse()
    Dim client As New HttpClient
    Dim Query As New HttpParams
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    client.DefaultHeaders.SetValue "X-Test-Header", "present"
    Query.Add "name", "value"

    Set response = client.GetResponse("/headers", Query)

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertTrue response.IsSuccess
    XlflowAssert.AssertContains "application/json", response.Headers.GetValue("Content-Type")
    XlflowAssert.AssertContains "present", response.Text
    XlflowAssert.AssertContains "value", response.Text
    XlflowAssert.AssertTrue response.ElapsedMilliseconds >= 0
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_SendsBufferedMethods()
    Dim client As New HttpClient
    Dim response As HttpResponse

    AssertTextEcho "POST"
    AssertTextEcho "PUT"
    AssertTextEcho "PATCH"

    client.BaseUrl = RequireBaseUrl()
    Set response = client.DeleteResponse("/echo")
    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertTrue response.Body.IsEmpty
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_PreservesBinaryBody()
    Dim client As New HttpClient
    Dim Request As New HttpRequest
    Dim Body As New HttpBody
    Dim response As HttpResponse
    Dim bytes(0 To 2) As Byte
    Dim responseBytes As Variant

    bytes(0) = 0
    bytes(1) = 127
    bytes(2) = 255
    Body.SetBytes bytes
    Request.Method = "POST"
    Request.Url = RequireBaseUrl() & "/echo"
    Request.Headers.SetValue "Content-Type", "application/octet-stream"
    Set Request.Body = Body

    Set response = client.Execute(Request)
    responseBytes = response.Body.Bytes

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals 3, HttpEncoding.ByteCount(responseBytes)
    XlflowAssert.AssertEquals 0, HttpEncoding.ByteAt(responseBytes, 0)
    XlflowAssert.AssertEquals 127, HttpEncoding.ByteAt(responseBytes, 1)
    XlflowAssert.AssertEquals 255, HttpEncoding.ByteAt(responseBytes, 2)
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_AppliesRedirectPolicy()
    Dim client As New HttpClient
    Dim Request As New HttpRequest
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set response = client.GetResponse("/redirect/2")
    XlflowAssert.AssertEquals 200, response.StatusCode

    Request.Url = "/redirect/2"
    Request.FollowRedirects = False
    Set response = client.Execute(Request)
    XlflowAssert.AssertEquals 302, response.StatusCode
    XlflowAssert.AssertEquals "/redirect/1", response.Headers.GetValue("Location")
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_ReturnsHttpFailureStatus()
    Dim client As New HttpClient
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set response = client.GetResponse("/status/503")

    XlflowAssert.AssertEquals 503, response.StatusCode
    XlflowAssert.AssertFalse response.IsSuccess
End Sub

'@Tag("integration")
'@ExpectedError(-2147200498, "WinHTTP request timed out (12002).", "WinHttpComTransport.Execute")
Public Sub Test_ComTransport_MapsReceiveTimeout()
    Dim client As New HttpClient
    Dim Request As New HttpRequest

    Request.Url = RequireBaseUrl() & "/delay/10000"
    Request.Timeouts.ResolveMilliseconds = 1000
    Request.Timeouts.ConnectMilliseconds = 1000
    Request.Timeouts.SendMilliseconds = 1000
    Request.Timeouts.ReceiveMilliseconds = 1000

    Call client.Execute(Request)
End Sub

'@Tag("integration")
'@ExpectedError(-2147200500, "WinHTTP could not establish or maintain the connection (12029).", "WinHttpComTransport.Execute")
Public Sub Test_ComTransport_MapsLoopbackConnectionFailure()
    Dim client As New HttpClient
    Dim Request As New HttpRequest

    Call RequireBaseUrl

    Request.Url = "http://127.0.0.1:1/status/200"
    Request.Timeouts.ResolveMilliseconds = 5000
    Request.Timeouts.ConnectMilliseconds = 5000
    Request.Timeouts.SendMilliseconds = 5000
    Request.Timeouts.ReceiveMilliseconds = 5000

    Call client.Execute(Request)
End Sub

Private Sub AssertTextEcho(ByVal Method As String)
    Dim client As New HttpClient
    Dim Request As New HttpRequest
    Dim Body As New HttpBody
    Dim response As HttpResponse

    Body.Text = "payload-" & LCase$(Method)
    Request.Method = Method
    Request.Url = RequireBaseUrl() & "/echo"
    Set Request.Body = Body

    Set response = client.Execute(Request)

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals Body.Text, response.Text
End Sub

Private Function RequireBaseUrl() As String
    RequireBaseUrl = Trim$(Environ$("VBA_HTTP_TEST_BASE_URL"))
    If Len(RequireBaseUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_BASE_URL is not set; run task test:integration."
        Exit Function
    End If
End Function
