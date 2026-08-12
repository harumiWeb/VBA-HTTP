Attribute VB_Name = "HttpClientTests"
Option Explicit

Public Sub Test_Client_ExecutesSnapshotThroughTransport()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim Request As New HttpRequest
    Dim response As HttpResponse

    configuredResponse.Initialize 200, "OK"
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    client.BaseUrl = "https://api.example.test/v1/"
    Request.Url = "/users"

    Set response = client.Execute(Request)

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals 1, transport.ExecuteCount
    XlflowAssert.AssertEquals "https://api.example.test/v1/users", transport.LastRequest.Url
    XlflowAssert.AssertEquals "/users", Request.Url
End Sub

Public Sub Test_Client_MergesQueryAndOverridesDefaultHeaders()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim Request As New HttpRequest

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    client.DefaultHeaders.SetValue "Accept", "application/json"
    client.DefaultHeaders.SetValue "X-Default", "kept"
    Request.Url = "https://example.test/items?fixed=1"
    Request.Headers.SetValue "accept", "text/plain"
    Request.Query.Add "q", "a b"

    Call client.Execute(Request)

    XlflowAssert.AssertEquals "https://example.test/items?fixed=1&q=a%20b", transport.LastRequest.Url
    XlflowAssert.AssertEquals "text/plain", transport.LastRequest.Headers.GetValue("Accept")
    XlflowAssert.AssertEquals "kept", transport.LastRequest.Headers.GetValue("X-Default")
    XlflowAssert.AssertEquals "application/json", client.DefaultHeaders.GetValue("Accept")
End Sub

Public Sub Test_Client_ResponseIsIndependentFromMockTemplate()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim response As HttpResponse

    configuredResponse.Initialize 200
    configuredResponse.Headers.SetValue "X-Template", "original"
    transport.SetResponse configuredResponse
    Set client.Transport = transport

    Set response = client.GetResponse("https://example.test/")
    response.Headers.SetValue "X-Template", "changed"

    Set response = client.GetResponse("https://example.test/")
    XlflowAssert.AssertEquals "original", response.Headers.GetValue("X-Template")
End Sub

Public Sub Test_Client_PropagatesTransportFailure()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim observedNumber As Long

    transport.SetFailure HttpErrTimeout, "timed out"
    Set client.Transport = transport
    On Error Resume Next
    Call client.GetResponse("https://example.test/")
    observedNumber = Err.Number
    Err.Clear
    On Error GoTo 0

    XlflowAssert.AssertEquals HttpErrTimeout, observedNumber
    XlflowAssert.AssertEquals HttpErrorTimeout, HttpErrors.CategoryFromNumber(observedNumber)
End Sub

Public Sub Test_Client_ConvenienceMethodsSetMethodAndBody()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim Body As New HttpBody

    configuredResponse.Initialize 201
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    Body.Text = "payload"

    Call client.PostResponse("https://example.test/items", Body)

    XlflowAssert.AssertEquals "POST", transport.LastRequest.Method
    XlflowAssert.AssertEquals "payload", transport.LastRequest.Body.Text
End Sub

Public Sub Test_Client_PropagatesDefaultProtocolOptionsToSnapshot()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim options As New HttpProtocolOptions

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    options.AllowHttp2 = True
    Set client.ProtocolOptions = options

    Call client.GetResponse("https://example.test/items")

    XlflowAssert.AssertTrue transport.LastRequest.ProtocolOptions.AllowHttp2
    XlflowAssert.AssertNotSame options, transport.LastRequest.ProtocolOptions
End Sub

Public Sub Test_Client_PropagatesDefaultDecompressionOptionsToSnapshot()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim options As New HttpDecompressionOptions

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    options.AllowGzip = True
    Set client.DecompressionOptions = options

    Call client.GetResponse("https://example.test/items")

    XlflowAssert.AssertTrue transport.LastRequest.DecompressionOptions.AllowGzip
    XlflowAssert.AssertNotSame options, transport.LastRequest.DecompressionOptions
End Sub

Public Sub Test_Client_PropagatesDefaultProxyOptionsToSnapshot()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim options As New HttpProxyOptions

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    options.Mode = HttpProxyNoProxy
    Set client.ProxyOptions = options

    Call client.GetResponse("https://example.test/items")

    XlflowAssert.AssertEquals HttpProxyNoProxy, transport.LastRequest.ProxyOptions.Mode
    XlflowAssert.AssertNotSame options, transport.LastRequest.ProxyOptions
End Sub

Public Sub Test_Client_DisablesRedirectsForSensitiveHeaders()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim Request As New HttpRequest
    Dim headerName As Variant

    configuredResponse.Initialize 302
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    Request.Url = "https://example.test/redirect"
    For Each headerName In Array("Authorization", "Proxy-Authorization", "Cookie")
        Request.Headers.Clear
        Request.Headers.SetValue CStr(headerName), "redacted-test-value"
        Request.FollowRedirects = True
        Call client.Execute(Request)
        XlflowAssert.AssertFalse transport.LastRequest.FollowRedirects, CStr(headerName) & " must suppress automatic redirects."
        XlflowAssert.AssertTrue Request.FollowRedirects, "The caller-owned request must not be mutated."
    Next headerName
End Sub

'@ExpectedError(-2147200502, "URL must contain a host.", "HttpClient.Execute")
Public Sub Test_Client_RejectsAbsoluteUrlWithoutHost()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport

    Set client.Transport = transport
    Call client.GetResponse("https://")
End Sub

'@ExpectedError(-2147200502, "Request URL must use HTTP or HTTPS.", "HttpClient.Execute")
Public Sub Test_Client_RejectsNonHttpScheme()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport

    client.BaseUrl = "https://example.test"
    Set client.Transport = transport
    Call client.GetResponse("ftp://example.test/file")
End Sub

Public Sub Test_Client_DefaultsToComTransport()
    Dim client As New HttpClient

    XlflowAssert.AssertEquals "WinHttpComTransport", TypeName(client.Transport)
End Sub

'@ExpectedError(-2147200503, "HTTP/2 and HTTP/3 protocol options require the native WinHTTP transport.", "WinHttpComTransport.Execute")
Public Sub Test_Client_ComTransportRejectsAdvancedProtocolOptions()
    Dim client As New HttpClient
    Dim options As New HttpProtocolOptions

    options.AllowHttp2 = True
    Set client.ProtocolOptions = options
    Call client.GetResponse("https://example.test/")
End Sub

'@ExpectedError(-2147200503, "Response decompression options require the native WinHTTP transport.", "WinHttpComTransport.Execute")
Public Sub Test_Client_ComTransportRejectsDecompressionOptions()
    Dim client As New HttpClient
    Dim options As New HttpDecompressionOptions

    options.AllowGzip = True
    Set client.DecompressionOptions = options
    Call client.GetResponse("https://example.test/")
End Sub
