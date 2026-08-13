Attribute VB_Name = "ReleaseBatchSmoke"
Option Explicit

' Keep this external harness independent from the production workbook's
' private modules. These values mirror HttpErrors.bas without creating a
' compile-time dependency on the filtered artifact.
Private Const ReleaseHttpErrTls As Long = vbObjectError + 21005
Private Const ReleaseHttpErrProtocol As Long = vbObjectError + 21008

Public Sub RunBatchSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim Urls As New Collection
    Dim result As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    client.BaseUrl = baseUrl
    Urls.Add "/delay/100"
    Urls.Add "/delay/100"
    Urls.Add "/delay/100"
    Urls.Add "/delay/100"
    Set result = client.GetMany(Urls)

    If result.Count <> 4 Or result.SuccessCount <> 4 Or result.FailureCount <> 0 Or result.CancelledCount <> 0 Then
        Err.Raise vbObjectError + 740, "ReleaseBatchSmoke.RunBatchSmoke", "Release batch smoke returned an unexpected result."
    End If
End Sub

Public Sub RunReliabilitySmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim policy As Object
    Dim Options As Object
    Dim response As Object
    Dim errorNumber As Long

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    Set policy = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateRetryPolicy")
    policy.BaseDelayMilliseconds = 0
    policy.MaxDelayMilliseconds = 0
    Set client.RetryPolicy = policy
    client.BaseUrl = baseUrl

    Set response = client.GetResponse("/flaky/2?id=release-reliability")
    If response.StatusCode <> 200 Or InStr(1, response.Text, """attempt"":3", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 741, "ReleaseBatchSmoke.RunReliabilitySmoke", "Release retry smoke returned an unexpected result."
    End If

    Set Options = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateExecutionOptions")
    Options.TotalDeadlineMilliseconds = 100
    Options.YieldToHost = False
    On Error Resume Next
    Set response = client.GetResponse("/delay/2000", Nothing, Options)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    If errorNumber <> -2147200498 Then
        Err.Raise vbObjectError + 742, "ReleaseBatchSmoke.RunReliabilitySmoke", "Release total-deadline smoke did not return HttpErrTimeout."
    End If
End Sub

Public Sub RunProtocolSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim protocols As Object
    Dim response As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateNativeClient")
    Set protocols = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateProtocolOptions")
    protocols.AllowHttp2 = True
    protocols.AllowHttp3 = True
    protocols.Mode = 0
    Set client.ProtocolOptions = protocols
    client.BaseUrl = baseUrl

    Set response = client.GetResponse("/status/204")
    If response.StatusCode <> 204 Or Not response.IsSuccess Or response.ProtocolUsed <> "HTTP/1.1" Then
        Err.Raise vbObjectError + 747, "ReleaseBatchSmoke.RunProtocolSmoke", "Release protocol fallback smoke returned an unexpected result."
    End If
End Sub

Public Function RunProtocolHostSmoke(ByVal releaseWorkbookName As String, ByVal Url As String, ByVal ExpectedProtocol As String) As String
    Dim client As Object
    Dim request As Object
    Dim protocols As Object
    Dim timeouts As Object
    Dim response As Object
    Dim expected As String

    expected = UCase$(Trim$(ExpectedProtocol))
    If expected <> "HTTP/2" And expected <> "HTTP/3" Then
        Err.Raise vbObjectError + 755, "ReleaseBatchSmoke.RunProtocolHostSmoke", "Expected protocol must be HTTP/2 or HTTP/3."
    End If

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateNativeClient")
    Set request = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateRequest")
    request.Method = "GET"
    request.Url = Url
    Set timeouts = request.Timeouts
    timeouts.ResolveMilliseconds = 15000
    timeouts.ConnectMilliseconds = 15000
    timeouts.SendMilliseconds = 15000
    timeouts.ReceiveMilliseconds = 15000
    Set protocols = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateProtocolOptions")
    If expected = "HTTP/2" Then
        protocols.AllowHttp2 = True
    Else
        protocols.AllowHttp3 = True
    End If
    protocols.Mode = 1
    Set request.ProtocolOptions = protocols
    Set client.ProtocolOptions = protocols
    Set response = client.Execute(request)

    If response.StatusCode < 200 Or response.StatusCode >= 300 Or response.ProtocolUsed <> expected Then
        Err.Raise vbObjectError + 756, "ReleaseBatchSmoke.RunProtocolHostSmoke", "Required protocol was not negotiated."
    End If
    RunProtocolHostSmoke = response.ProtocolUsed
End Function

Public Sub RunDecompressionSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim options As Object
    Dim response As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateNativeClient")
    Set options = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateDecompressionOptions")
    options.AllowGzip = True
    Set client.DecompressionOptions = options
    client.BaseUrl = baseUrl

    Set response = client.GetResponse("/compress/gzip")
    If response.StatusCode <> 200 Or response.Text <> "VBA-HTTP compression fixture: 0123456789" & vbLf Then
        Err.Raise vbObjectError + 748, "ReleaseBatchSmoke.RunDecompressionSmoke", "Release decompression smoke returned an unexpected response."
    End If
End Sub

Public Sub RunProxySmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String, ByVal proxyUrl As String, ByVal proxyAuthUrl As String, ByVal proxyTlsTargetUrl As String, ByVal proxyTlsUrl As String, ByVal proxyTlsAuthUrl As String)
    Dim client As Object
    Dim nativeClient As Object
    Dim options As Object
    Dim provider As Object
    Dim response As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    Set options = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateProxyOptions")
    options.Mode = 2
    options.ProxyUrl = proxyUrl
    Set client.ProxyOptions = options
    client.BaseUrl = baseUrl
    Set response = client.GetResponse("/headers")
    If response.StatusCode <> 200 Or response.Headers.GetValue("X-Test-Proxy-Forwarded") <> "1" Then
        Err.Raise vbObjectError + 749, "ReleaseBatchSmoke.RunProxySmoke", "Release COM proxy smoke returned an unexpected response."
    End If

    Set nativeClient = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateNativeClient")
    Set options = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateProxyOptions")
    options.Mode = 2
    options.ProxyUrl = proxyUrl
    Set nativeClient.ProxyOptions = options
    nativeClient.BaseUrl = baseUrl
    Set response = nativeClient.GetResponse("/headers")
    If response.StatusCode <> 200 Or response.Headers.GetValue("X-Test-Proxy-Forwarded") <> "1" Then
        Err.Raise vbObjectError + 750, "ReleaseBatchSmoke.RunProxySmoke", "Release native proxy smoke returned an unexpected response."
    End If

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    Set options = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateProxyOptions")
    options.Mode = 2
    options.ProxyUrl = proxyAuthUrl
    Set client.ProxyOptions = options
    Set provider = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateWindowsAuthProvider", "proxy-user", "proxy-pass", 1, 1, True, 3)
    Set client.AuthProvider = provider
    client.BaseUrl = baseUrl
    Set response = client.GetResponse("/headers")
    If response.StatusCode <> 200 Or response.Headers.GetValue("X-Test-Proxy-Forwarded") <> "1" Then
        Err.Raise vbObjectError + 753, "ReleaseBatchSmoke.RunProxySmoke", "Release COM proxy challenge smoke returned an unexpected response."
    End If

    Set nativeClient = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateNativeClient")
    Set options = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateProxyOptions")
    options.Mode = 2
    options.ProxyUrl = proxyAuthUrl
    Set nativeClient.ProxyOptions = options
    Set provider = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateWindowsAuthProvider", "proxy-user", "proxy-pass", 1, 1, True, 3)
    Set nativeClient.AuthProvider = provider
    nativeClient.BaseUrl = baseUrl
    Set response = nativeClient.GetResponse("/headers")
    If response.StatusCode <> 200 Or response.Headers.GetValue("X-Test-Proxy-Forwarded") <> "1" Then
        Err.Raise vbObjectError + 754, "ReleaseBatchSmoke.RunProxySmoke", "Release native proxy challenge smoke returned an unexpected response."
    End If

    AssertProxyTlsBoundary releaseWorkbookName, baseUrl, proxyTlsTargetUrl, proxyTlsUrl, False, False
    AssertProxyTlsBoundary releaseWorkbookName, baseUrl, proxyTlsTargetUrl, proxyTlsAuthUrl, True, False
    AssertProxyTlsBoundary releaseWorkbookName, baseUrl, proxyTlsTargetUrl, proxyTlsUrl, False, True
    AssertProxyTlsBoundary releaseWorkbookName, baseUrl, proxyTlsTargetUrl, proxyTlsAuthUrl, True, True
End Sub

Private Sub AssertProxyTlsBoundary(ByVal releaseWorkbookName As String, ByVal baseUrl As String, ByVal targetUrl As String, ByVal proxyUrl As String, ByVal authenticated As Boolean, ByVal nativeTransport As Boolean)
    Dim statsClient As Object
    Dim client As Object
    Dim options As Object
    Dim provider As Object
    Dim beforeResponse As Object
    Dim afterResponse As Object
    Dim response As Object
    Dim beforeAttempts As Long
    Dim beforeAuthorized As Long
    Dim observedNumber As Long
    Dim transportName As String

    If nativeTransport Then
        transportName = "native"
        Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateNativeClient")
    Else
        transportName = "COM"
        Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    End If
    Set statsClient = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    statsClient.BaseUrl = baseUrl
    Set beforeResponse = statsClient.GetResponse("/__admin/proxy-stats")
    beforeAttempts = ExtractJsonLong(beforeResponse.Text, "connect_attempts")
    beforeAuthorized = ExtractJsonLong(beforeResponse.Text, "authorized_connects")

    Set options = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateProxyOptions")
    options.Mode = 2
    options.ProxyUrl = proxyUrl
    Set client.ProxyOptions = options
    client.BaseUrl = targetUrl
    If authenticated Then
        Set provider = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateWindowsAuthProvider", "proxy-user", "proxy-pass", 1, 1, True, 3)
        Set client.AuthProvider = provider
    End If

    On Error Resume Next
    Set response = client.GetResponse("/status/204")
    observedNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    If observedNumber = 0 Or observedNumber <> ReleaseHttpErrTls Then
        Err.Raise vbObjectError + 763, "ReleaseBatchSmoke.RunProxySmoke", "Release " & transportName & " HTTPS CONNECT boundary did not reject the untrusted TLS certificate."
    End If

    Set afterResponse = statsClient.GetResponse("/__admin/proxy-stats")
    If ExtractJsonLong(afterResponse.Text, "connect_attempts") <= beforeAttempts Then
        Err.Raise vbObjectError + 764, "ReleaseBatchSmoke.RunProxySmoke", "Release " & transportName & " HTTPS CONNECT boundary did not reach the proxy."
    End If
    If authenticated And ExtractJsonLong(afterResponse.Text, "authorized_connects") <= beforeAuthorized Then
        Err.Raise vbObjectError + 765, "ReleaseBatchSmoke.RunProxySmoke", "Release " & transportName & " authenticated HTTPS CONNECT boundary did not authorize the proxy."
    End If
End Sub

Private Function ExtractJsonLong(ByVal Text As String, ByVal Name As String) As Long
    Dim marker As String
    Dim startIndex As Long
    Dim endIndex As Long

    marker = Chr$(34) & Name & Chr$(34) & ":"
    startIndex = InStr(1, Text, marker, vbBinaryCompare)
    If startIndex = 0 Then Err.Raise ReleaseHttpErrProtocol, "ReleaseBatchSmoke", "JSON field was not found: " & Name
    startIndex = startIndex + Len(marker)
    Do While startIndex <= Len(Text) And Mid$(Text, startIndex, 1) = " "
        startIndex = startIndex + 1
    Loop
    endIndex = startIndex
    Do While endIndex <= Len(Text) And Mid$(Text, endIndex, 1) Like "[0-9]"
        endIndex = endIndex + 1
    Loop
    If endIndex <= startIndex Then Err.Raise ReleaseHttpErrProtocol, "ReleaseBatchSmoke", "JSON field was malformed: " & Name
    ExtractJsonLong = CLng(Mid$(Text, startIndex, endIndex - startIndex))
End Function

Public Sub RunAuthSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim nativeClient As Object
    Dim provider As Object
    Dim response As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    client.BaseUrl = baseUrl
    Set provider = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateBasicAuthProvider", "user", "pass", True)
    Set client.AuthProvider = provider
    Set response = client.GetResponse("/auth/basic")
    If response.StatusCode <> 204 Or response.Headers.GetValue("X-Auth-Verified") <> "1" Then
        Err.Raise vbObjectError + 751, "ReleaseBatchSmoke.RunAuthSmoke", "Release Basic authentication smoke returned an unexpected response."
    End If

    Set provider = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateBearerAuthProvider", "vba-http-token", True)
    Set client.AuthProvider = provider
    Set response = client.GetResponse("/auth/bearer")
    If response.StatusCode <> 204 Or response.Headers.GetValue("X-Auth-Verified") <> "1" Then
        Err.Raise vbObjectError + 752, "ReleaseBatchSmoke.RunAuthSmoke", "Release Bearer authentication smoke returned an unexpected response."
    End If

    Set provider = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateWindowsAuthProvider", "user", "pass", 1, 0, True, 3)
    Set client.AuthProvider = provider
    Set response = client.GetResponse("/auth/challenge/basic")
    If response.StatusCode <> 204 Or response.Headers.GetValue("X-Auth-Verified") <> "1" Then
        Err.Raise vbObjectError + 761, "ReleaseBatchSmoke.RunAuthSmoke", "Release COM challenge authentication smoke returned an unexpected response."
    End If

    Set nativeClient = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateNativeClient")
    nativeClient.BaseUrl = baseUrl
    Set provider = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateWindowsAuthProvider", "user", "pass", 1, 0, True, 3)
    Set nativeClient.AuthProvider = provider
    Set response = nativeClient.GetResponse("/auth/challenge/basic")
    If response.StatusCode <> 204 Or response.Headers.GetValue("X-Auth-Verified") <> "1" Then
        Err.Raise vbObjectError + 762, "ReleaseBatchSmoke.RunAuthSmoke", "Release native challenge authentication smoke returned an unexpected response."
    End If
End Sub

Public Sub RunDiagnosticsSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim diagnostics As Object
    Dim response As Object
    Dim json As String

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    Set diagnostics = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateDiagnostics")
    diagnostics.Enabled = True
    Set client.Diagnostics = diagnostics
    client.BaseUrl = baseUrl
    client.DefaultHeaders.SetValue "Authorization", "release-diagnostics-secret"
    client.DefaultHeaders.SetValue "Cookie", "release=session-secret"
    Set response = client.GetResponse("/status/204?query-secret=release-query-secret")
    json = diagnostics.ToJson
    If response.StatusCode <> 204 Or diagnostics.Count <> 1 Then
        Err.Raise vbObjectError + 759, "ReleaseBatchSmoke.RunDiagnosticsSmoke", "Release diagnostics smoke returned an unexpected event count."
    End If
    If InStr(1, json, "[REDACTED]", vbBinaryCompare) = 0 Or _
        InStr(1, json, "release-diagnostics-secret", vbBinaryCompare) > 0 Or _
        InStr(1, json, "session-secret", vbBinaryCompare) > 0 Or _
        InStr(1, json, "release-query-secret", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 760, "ReleaseBatchSmoke.RunDiagnosticsSmoke", "Release diagnostics smoke exposed a sensitive value."
    End If
End Sub

Public Sub RunCookieSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim jar As Object
    Dim response As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    Set jar = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateCookieJar")
    Set client.CookieJar = jar
    client.BaseUrl = baseUrl

    Set response = client.GetResponse("/cookie/set")
    If response.StatusCode <> 204 Or jar.Count <> 1 Then
        Err.Raise vbObjectError + 755, "ReleaseBatchSmoke.RunCookieSmoke", "Release cookie jar did not store the fixture cookie."
    End If

    Set response = client.GetResponse("/cookie/echo")
    If response.StatusCode <> 204 Or response.Headers.GetValue("X-Cookie-Verified") <> "1" Then
        Err.Raise vbObjectError + 756, "ReleaseBatchSmoke.RunCookieSmoke", "Release cookie jar did not send the fixture cookie."
    End If

    Set response = client.GetResponse("/cookie/redirect")
    If response.StatusCode <> 302 Or Len(response.Headers.GetValue("Location")) = 0 Then
        Err.Raise vbObjectError + 757, "ReleaseBatchSmoke.RunCookieSmoke", "Release cookie jar followed a credential-bearing redirect."
    End If

    Set response = client.GetResponse("/cookie/clear")
    If response.StatusCode <> 204 Or jar.Count <> 0 Then
        Err.Raise vbObjectError + 758, "ReleaseBatchSmoke.RunCookieSmoke", "Release cookie jar did not clear the fixture cookie."
    End If
End Sub

Public Sub RunRedirectSecuritySmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim nativeClient As Object
    Dim response As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    client.BaseUrl = baseUrl
    client.DefaultHeaders.SetValue "Authorization", "redacted-release-smoke"
    Set response = client.GetResponse("/redirect/1")
    If response.StatusCode <> 302 Then
        Err.Raise vbObjectError + 753, "ReleaseBatchSmoke.RunRedirectSecuritySmoke", "Release COM redirect security smoke followed a credential-bearing redirect."
    End If

    Set nativeClient = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateNativeClient")
    nativeClient.BaseUrl = baseUrl
    nativeClient.DefaultHeaders.SetValue "Cookie", "redacted-release-smoke"
    Set response = nativeClient.GetResponse("/redirect/1")
    If response.StatusCode <> 302 Then
        Err.Raise vbObjectError + 754, "ReleaseBatchSmoke.RunRedirectSecuritySmoke", "Release native redirect security smoke followed a cookie-bearing redirect."
    End If
End Sub

Public Sub RunUploadSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String, ByVal sourcePath As String)
    Dim client As Object
    Dim form As Object
    Dim result As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateNativeClient")
    client.BaseUrl = baseUrl

    Set result = client.UploadFile("/upload/hash", sourcePath, "application/octet-stream")
    If result.StatusCode <> 200 Or result.BytesWritten <> 65536 Or result.ContentLength <> 65536 Then
        Err.Raise vbObjectError + 743, "ReleaseBatchSmoke.RunUploadSmoke", "Release file upload smoke returned an unexpected result."
    End If
    If result.Headers.GetValue("X-Upload-Bytes") <> "65536" Then
        Err.Raise vbObjectError + 744, "ReleaseBatchSmoke.RunUploadSmoke", "Release file upload byte header was not verified."
    End If

    Set form = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateMultipartForm")
    form.Boundary = "----vba-http-release-smoke"
    form.AddField "title", "外部consumer"
    form.AddFile "payload", sourcePath, "consumer.bin", "application/octet-stream"
    Set result = client.UploadMultipart("/upload/multipart", form)
    If result.StatusCode <> 200 Or result.BytesWritten <> result.ContentLength Then
        Err.Raise vbObjectError + 745, "ReleaseBatchSmoke.RunUploadSmoke", "Release multipart upload smoke returned an unexpected result."
    End If
    If result.Headers.GetValue("X-Multipart-File-Bytes") <> "65536" Or result.Headers.GetValue("X-Multipart-Filename") <> "consumer.bin" Then
        Err.Raise vbObjectError + 746, "ReleaseBatchSmoke.RunUploadSmoke", "Release multipart response was not verified."
    End If
End Sub
