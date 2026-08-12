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
Public Sub Test_ComTransport_ReturnsUnicodeResponseText()
    Dim client As New HttpClient
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set response = client.GetResponse("/unicode")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals UnicodeFixtureText(), response.Text
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_ReturnsDirectBinaryResponse()
    Dim client As New HttpClient
    Dim response As HttpResponse
    Dim responseBytes As Variant

    client.BaseUrl = RequireBaseUrl()
    Set response = client.GetResponse("/bytes/3")
    responseBytes = response.Body.Bytes

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals 3, HttpEncoding.ByteCount(responseBytes)
    XlflowAssert.AssertEquals 0, HttpEncoding.ByteAt(responseBytes, 0)
    XlflowAssert.AssertEquals 1, HttpEncoding.ByteAt(responseBytes, 1)
    XlflowAssert.AssertEquals 2, HttpEncoding.ByteAt(responseBytes, 2)
End Sub

'@Tag("integration")
'@ExpectedError(-2147200503, "Response contains malformed UTF-8.", "HttpEncoding.DecodeUtf8")
Public Sub Test_ComTransport_RejectsMalformedUtf8ResponseText()
    Dim client As New HttpClient
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set response = client.GetResponse("/malformed-utf8")
    Debug.Print response.Text
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_MapsMalformedResponseHeadersToProtocolError()
    Dim client As New HttpClient
    Dim observedNumber As Long

    client.BaseUrl = RequireBaseUrl()
    On Error GoTo ExpectedFailure
    Call client.GetResponse("/malformed-headers")
    XlflowAssert.AssertTrue False, "Malformed response headers should fail the exchange."
    Exit Sub

    ExpectedFailure: ' xlflow:disable-line VBA237
    observedNumber = Err.Number
    Err.Clear
    XlflowAssert.AssertEquals HttpErrorProtocol, HttpErrors.CategoryFromNumber(observedNumber)
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
Public Sub Test_ComTransport_BasicAndBearerAuth()
    Dim client As New HttpClient
    Dim provider As IHttpAuthProvider
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set provider = VBAHttp.CreateBasicAuthProvider("user", "pass", True)
    Set client.AuthProvider = provider
    Set response = client.GetResponse("/auth/basic")
    XlflowAssert.AssertEquals 204, response.StatusCode
    XlflowAssert.AssertEquals "1", response.Headers.GetValue("X-Auth-Verified")

    Set provider = VBAHttp.CreateBasicAuthProvider("user", "wrong", True)
    Set client.AuthProvider = provider
    Set response = client.GetResponse("/auth/basic")
    XlflowAssert.AssertEquals 401, response.StatusCode

    Set provider = VBAHttp.CreateBearerAuthProvider("vba-http-token", True)
    Set client.AuthProvider = provider
    Set response = client.GetResponse("/auth/bearer")
    XlflowAssert.AssertEquals 204, response.StatusCode
    XlflowAssert.AssertEquals "1", response.Headers.GetValue("X-Auth-Verified")
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_AuthProviderDisablesRedirects()
    Dim client As New HttpClient
    Dim provider As IHttpAuthProvider
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set provider = VBAHttp.CreateBasicAuthProvider("user", "pass", True)
    Set client.AuthProvider = provider
    Set response = client.GetResponse("/redirect/1")

    XlflowAssert.AssertEquals 302, response.StatusCode
    XlflowAssert.AssertEquals "/redirect/0", response.Headers.GetValue("Location")
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_SensitiveHeadersDisableRedirects()
    Dim client As New HttpClient
    Dim Request As New HttpRequest
    Dim response As HttpResponse
    Dim headerName As Variant

    client.BaseUrl = RequireBaseUrl()
    Request.Url = "/redirect/1"
    For Each headerName In Array("Authorization", "Proxy-Authorization", "Cookie")
        Request.Headers.Clear
        Request.Headers.SetValue CStr(headerName), "redacted-test-value"
        Request.FollowRedirects = True
        Set response = client.Execute(Request)
        XlflowAssert.AssertEquals 302, response.StatusCode
        XlflowAssert.AssertEquals "/redirect/0", response.Headers.GetValue("Location")
    Next headerName
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_BoundsRedirectLoop()
    Dim client As New HttpClient
    Dim Request As New HttpRequest
    Dim observedNumber As Long

    Request.Url = RequireBaseUrl() & "/redirect-loop"
    Request.MaxRedirects = 2
    On Error GoTo ExpectedFailure
    Call client.Execute(Request)
    XlflowAssert.AssertTrue False, "Redirect loop should exceed the configured maximum."
    Exit Sub

    ExpectedFailure: ' xlflow:disable-line VBA237
    observedNumber = Err.Number
    Err.Clear
    XlflowAssert.AssertEquals HttpErrorProtocol, HttpErrors.CategoryFromNumber(observedNumber)
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
Public Sub Test_ComTransport_UsesManualProxy()
    Dim client As New HttpClient
    Dim options As New HttpProxyOptions
    Dim response As HttpResponse

    client.BaseUrl = RequireProxyTargetUrl()
    options.Mode = HttpProxyManual
    options.ProxyUrl = RequireProxyUrl()
    Set client.ProxyOptions = options

    Set response = client.GetResponse("/headers")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals "1", response.Headers.GetValue("X-Test-Proxy-Forwarded")
    XlflowAssert.AssertContains "X-Test-Proxy-Forwarded", response.Text
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

Private Function UnicodeFixtureText() As String
    UnicodeFixtureText = "VBA-HTTP unicode: " & ChrW$(&H65E5) & ChrW$(&H672C) & ChrW$(&H8A9E) & ChrW$(&HD83D) & ChrW$(&HDE42)
End Function

Private Function RequireProxyUrl() As String
    RequireProxyUrl = Trim$(Environ$("VBA_HTTP_TEST_PROXY_URL"))
    If Len(RequireProxyUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_PROXY_URL is not set; run task test:integration."
    End If
End Function

Private Function RequireProxyTargetUrl() As String
    RequireProxyTargetUrl = Trim$(Environ$("VBA_HTTP_TEST_PROXY_TARGET_URL"))
    If Len(RequireProxyTargetUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_PROXY_TARGET_URL is not set; run task test:integration."
    End If
End Function
