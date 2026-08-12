Attribute VB_Name = "ReleaseBatchSmoke"
Option Explicit

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

Public Sub RunProxySmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String, ByVal proxyUrl As String)
    Dim client As Object
    Dim nativeClient As Object
    Dim options As Object
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
End Sub

Public Sub RunAuthSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
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
