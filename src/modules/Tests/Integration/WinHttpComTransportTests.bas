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
Public Sub Test_ComTransport_BoundedChallengeAuth()
    Dim client As New HttpClient
    Dim provider As IHttpAuthProvider
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set provider = VBAHttp.CreateWindowsAuthProvider("user", "pass", HttpAuthSchemeAuto, HttpAuthTargetServer, True)
    Set client.AuthProvider = provider
    Set response = client.GetResponse("/auth/challenge/basic")

    XlflowAssert.AssertEquals 204, response.StatusCode
    XlflowAssert.AssertEquals "1", response.Headers.GetValue("X-Auth-Verified")
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_RejectsUntrustedCertificate()
    Dim client As New HttpClient
    Dim observedNumber As Long

    client.BaseUrl = RequireHttpsBaseUrl()
    On Error GoTo ExpectedFailure
    Call client.GetResponse("/status/204")
    XlflowAssert.AssertTrue False, "COM transport accepted the untrusted TLS fixture."
    Exit Sub

    ExpectedFailure: ' xlflow:disable-line VBA237
    observedNumber = Err.Number
    Err.Clear
    XlflowAssert.AssertEquals HttpErrorTls, HttpErrors.CategoryFromNumber(observedNumber)
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_ExplicitCookieJarPersistsAndSuppressesRedirects()
    Dim client As New HttpClient
    Dim jar As New HttpCookieJar
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set client.CookieJar = jar
    Set response = client.GetResponse("/cookie/set")
    XlflowAssert.AssertEquals 204, response.StatusCode
    XlflowAssert.AssertEquals 1, jar.Count

    Set response = client.GetResponse("/cookie/echo")
    XlflowAssert.AssertEquals 204, response.StatusCode
    XlflowAssert.AssertEquals "1", response.Headers.GetValue("X-Cookie-Verified")

    Set response = client.GetResponse("/cookie/redirect")
    XlflowAssert.AssertEquals 302, response.StatusCode
    XlflowAssert.AssertEquals "/cookie/echo", response.Headers.GetValue("Location")

    Set response = client.GetResponse("/cookie/clear")
    XlflowAssert.AssertEquals 204, response.StatusCode
    XlflowAssert.AssertEquals 0, jar.Count
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
Public Sub Test_ComTransport_ProxyChallengeAuth()
    Dim client As New HttpClient
    Dim options As New HttpProxyOptions
    Dim provider As IHttpAuthProvider
    Dim response As HttpResponse

    client.BaseUrl = RequireProxyTargetUrl()
    options.Mode = HttpProxyManual
    options.ProxyUrl = RequireProxyAuthUrl()
    Set client.ProxyOptions = options
    Set provider = VBAHttp.CreateWindowsAuthProvider("proxy-user", "proxy-pass", HttpAuthSchemeAuto, HttpAuthTargetProxy, True, 3)
    Set client.AuthProvider = provider

    Set response = client.GetResponse("/headers")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals "1", response.Headers.GetValue("X-Test-Proxy-Forwarded")
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_ProxyChallengeWrongCredentialsRemain407()
    Dim client As New HttpClient
    Dim options As New HttpProxyOptions
    Dim provider As IHttpAuthProvider
    Dim response As HttpResponse

    client.BaseUrl = RequireProxyTargetUrl()
    options.Mode = HttpProxyManual
    options.ProxyUrl = RequireProxyAuthUrl()
    Set client.ProxyOptions = options
    Set provider = VBAHttp.CreateWindowsAuthProvider("proxy-user", "wrong-pass", HttpAuthSchemeAuto, HttpAuthTargetProxy, True, 3)
    Set client.AuthProvider = provider

    Set response = client.GetResponse("/headers")

    XlflowAssert.AssertEquals 407, response.StatusCode
    XlflowAssert.AssertEquals "Basic realm=" & Chr$(34) & "vba-http-proxy-challenge" & Chr$(34), response.Headers.GetValue("Proxy-Authenticate")
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_HTTPSProxyConnectReachesTLSBoundary()
    Dim client As New HttpClient
    Dim statsClient As New HttpClient
    Dim options As New HttpProxyOptions
    Dim beforeResponse As HttpResponse
    Dim afterResponse As HttpResponse
    Dim observedNumber As Long
    Dim beforeAttempts As Long

    statsClient.BaseUrl = RequireBaseUrl()
    Set beforeResponse = statsClient.GetResponse("/__admin/proxy-stats")
    beforeAttempts = ExtractJsonLong(beforeResponse.Text, "connect_attempts")

    client.BaseUrl = RequireProxyTlsTargetUrl()
    options.Mode = HttpProxyManual
    options.ProxyUrl = RequireProxyTlsUrl()
    Set client.ProxyOptions = options
    On Error GoTo ExpectedFailure
    Call client.GetResponse("/status/204")
    XlflowAssert.AssertTrue False, "COM CONNECT proxy accepted the untrusted TLS fixture."
    Exit Sub

    ExpectedFailure: ' xlflow:disable-line VBA237
    observedNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    XlflowAssert.AssertEquals HttpErrorTls, HttpErrors.CategoryFromNumber(observedNumber)

    Set afterResponse = statsClient.GetResponse("/__admin/proxy-stats")
    XlflowAssert.AssertTrue ExtractJsonLong(afterResponse.Text, "connect_attempts") > beforeAttempts
End Sub

'@Tag("integration")
Public Sub Test_ComTransport_AuthenticatedHTTPSProxyConnectReachesTLSBoundary()
    Dim client As New HttpClient
    Dim statsClient As New HttpClient
    Dim options As New HttpProxyOptions
    Dim provider As IHttpAuthProvider
    Dim beforeResponse As HttpResponse
    Dim afterResponse As HttpResponse
    Dim observedNumber As Long
    Dim beforeAuthorized As Long

    statsClient.BaseUrl = RequireBaseUrl()
    Set beforeResponse = statsClient.GetResponse("/__admin/proxy-stats")
    beforeAuthorized = ExtractJsonLong(beforeResponse.Text, "authorized_connects")

    client.BaseUrl = RequireProxyTlsTargetUrl()
    options.Mode = HttpProxyManual
    options.ProxyUrl = RequireProxyTlsAuthUrl()
    Set client.ProxyOptions = options
    Set provider = VBAHttp.CreateWindowsAuthProvider("proxy-user", "proxy-pass", HttpAuthSchemeAuto, HttpAuthTargetProxy, True, 3)
    Set client.AuthProvider = provider
    On Error GoTo ExpectedFailure
    Call client.GetResponse("/status/204")
    XlflowAssert.AssertTrue False, "COM authenticated CONNECT proxy accepted the untrusted TLS fixture."
    Exit Sub

    ExpectedFailure: ' xlflow:disable-line VBA237
    observedNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    XlflowAssert.AssertEquals HttpErrorTls, HttpErrors.CategoryFromNumber(observedNumber)

    Set afterResponse = statsClient.GetResponse("/__admin/proxy-stats")
    XlflowAssert.AssertTrue ExtractJsonLong(afterResponse.Text, "authorized_connects") > beforeAuthorized
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

Private Function RequireProxyAuthUrl() As String
    RequireProxyAuthUrl = Trim$(Environ$("VBA_HTTP_TEST_PROXY_AUTH_URL"))
    If Len(RequireProxyAuthUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_PROXY_AUTH_URL is not set; run task test:integration."
    End If
End Function

Private Function RequireProxyTlsUrl() As String
    RequireProxyTlsUrl = Trim$(Environ$("VBA_HTTP_TEST_PROXY_TLS_URL"))
    If Len(RequireProxyTlsUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_PROXY_TLS_URL is not set; run task test:integration."
    End If
End Function

Private Function RequireProxyTlsTargetUrl() As String
    RequireProxyTlsTargetUrl = Trim$(Environ$("VBA_HTTP_TEST_PROXY_TLS_TARGET_URL"))
    If Len(RequireProxyTlsTargetUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_PROXY_TLS_TARGET_URL is not set; run task test:integration."
    End If
End Function

Private Function RequireProxyTlsAuthUrl() As String
    RequireProxyTlsAuthUrl = Trim$(Environ$("VBA_HTTP_TEST_PROXY_TLS_AUTH_URL"))
    If Len(RequireProxyTlsAuthUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_PROXY_TLS_AUTH_URL is not set; run task test:integration."
    End If
End Function

Private Function RequireHttpsBaseUrl() As String
    RequireHttpsBaseUrl = Trim$(Environ$("VBA_HTTP_TEST_HTTPS_URL"))
    If Len(RequireHttpsBaseUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_HTTPS_URL is not set; run task test:integration."
    End If
End Function

Private Function ExtractJsonLong(ByVal Text As String, ByVal Name As String) As Long
    Dim marker As String
    Dim startIndex As Long
    Dim endIndex As Long

    marker = Chr$(34) & Name & Chr$(34) & ":"
    startIndex = InStr(1, Text, marker, vbBinaryCompare)
    If startIndex = 0 Then Err.Raise HttpErrProtocol, "WinHttpComTransportTests", "JSON field was not found: " & Name
    startIndex = startIndex + Len(marker)
    Do While startIndex <= Len(Text) And Mid$(Text, startIndex, 1) = " "
        startIndex = startIndex + 1
    Loop
    endIndex = startIndex
    Do While endIndex <= Len(Text) And Mid$(Text, endIndex, 1) Like "[0-9]"
        endIndex = endIndex + 1
    Loop
    If endIndex <= startIndex Then Err.Raise HttpErrProtocol, "WinHttpComTransportTests", "JSON field was malformed: " & Name
    ExtractJsonLong = CLng(Mid$(Text, startIndex, endIndex - startIndex))
End Function
