Attribute VB_Name = "WinHttpNativeTransportTests"
Option Explicit

'@Tag("integration")
Public Sub Test_NativeTransport_GetMatchesComContract()
    Dim client As New HttpClient
    Dim Query As New HttpParams
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    client.DefaultHeaders.SetValue "X-Test-Header", "native-present"
    Query.Add "name", "native-value"

    Set client.Transport = New WinHttpNativeTransport
    Set response = client.GetResponse("/headers", Query)

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertTrue response.IsSuccess
    XlflowAssert.AssertContains "native-present", response.Text
    XlflowAssert.AssertContains "native-value", response.Text
    XlflowAssert.AssertTrue Len(response.ProtocolUsed) > 0
    XlflowAssert.AssertTrue response.ElapsedMilliseconds >= 0
End Sub

'@Tag("integration")
Public Sub Test_NativeTransport_SendsBufferedMethods()
    AssertTextEcho "POST"
    AssertTextEcho "PUT"
    AssertTextEcho "PATCH"
End Sub

'@Tag("integration")
Public Sub Test_NativeTransport_PreservesBinaryBody()
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
    Set client.Transport = New WinHttpNativeTransport

    Set response = client.Execute(Request)
    responseBytes = response.Body.Bytes

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals 3, HttpEncoding.ByteCount(responseBytes)
    XlflowAssert.AssertEquals 0, HttpEncoding.ByteAt(responseBytes, 0)
    XlflowAssert.AssertEquals 127, HttpEncoding.ByteAt(responseBytes, 1)
    XlflowAssert.AssertEquals 255, HttpEncoding.ByteAt(responseBytes, 2)
End Sub

'@Tag("integration")
Public Sub Test_NativeTransport_AppliesRedirectAndReturnsStatus()
    Dim client As New HttpClient
    Dim Request As New HttpRequest
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set client.Transport = New WinHttpNativeTransport
    Set response = client.GetResponse("/redirect/2")
    XlflowAssert.AssertEquals 200, response.StatusCode

    Request.Url = "/redirect/2"
    Request.FollowRedirects = False
    Set response = client.Execute(Request)
    XlflowAssert.AssertEquals 302, response.StatusCode
    XlflowAssert.AssertEquals "/redirect/1", response.Headers.GetValue("Location")
End Sub

'@Tag("integration")
Public Sub Test_NativeTransport_MapsConnectionFailure()
    Dim client As New HttpClient
    Dim Request As New HttpRequest
    Dim observedNumber As Long
    Dim observedSource As String
    Dim observedDescription As String

    Request.Url = "http://127.0.0.1:1/status/200"
    Request.Timeouts.ResolveMilliseconds = 5000
    Request.Timeouts.ConnectMilliseconds = 5000
    Request.Timeouts.SendMilliseconds = 5000
    Request.Timeouts.ReceiveMilliseconds = 5000
    Set client.Transport = New WinHttpNativeTransport

    On Error GoTo ExpectedFailure
    Call client.Execute(Request)
    XlflowAssert.AssertTrue False, "A loopback connection failure was expected."
    Exit Sub

    ' The expected-error branch intentionally handles and returns after the assertion.
    ExpectedFailure: ' xlflow:disable-line VBA237
    observedNumber = Err.Number
    observedSource = Err.Source
    observedDescription = Err.Description
    Err.Clear
    XlflowAssert.AssertEquals HttpErrorConnection, HttpErrors.CategoryFromNumber(observedNumber), "Native error: " & observedDescription
    XlflowAssert.AssertEquals "WinHttpNativeTransport.Execute", observedSource
End Sub

'@Tag("integration")
Public Sub Test_NativeTransport_ReleasesHandles()
    Dim client As New HttpClient
    Dim beforeCount As Long
    Dim afterCount As Long
    Dim index As Long

    If Not WinHttpNativeApi.ProcessHandleCount(beforeCount) Then
        XlflowAssert.AssertInconclusive "GetProcessHandleCount is unavailable on this host."
        Exit Sub
    End If
    client.BaseUrl = RequireBaseUrl()
    Set client.Transport = New WinHttpNativeTransport
    For index = 1 To 20
        Call client.GetResponse("/status/200")
    Next index
    If Not WinHttpNativeApi.ProcessHandleCount(afterCount) Then
        XlflowAssert.AssertInconclusive "GetProcessHandleCount is unavailable after requests."
        Exit Sub
    End If
    XlflowAssert.AssertTrue afterCount - beforeCount <= 8, "Native request handles grew persistently."
End Sub

'@Tag("integration")
Public Sub Test_NativeTransport_DownloadsKnownLengthAndReportsProgress()
    Dim client As New HttpClient
    Dim options As New HttpExecutionOptions
    Dim sink As New RecordingProgressSink
    Dim result As HttpDownloadResult
    Dim destination As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    destination = NewDownloadDestination("known")
    On Error GoTo Cleanup
    client.BaseUrl = RequireBaseUrl()
    Set client.Transport = New WinHttpNativeTransport
    Set result = client.DownloadFile("/bytes/1048576", destination, options, sink)

    XlflowAssert.AssertEquals 200, result.StatusCode
    XlflowAssert.AssertTrue result.IsSuccess
    XlflowAssert.AssertTrue result.Published
    XlflowAssert.AssertEquals CCur(1048576), result.BytesWritten
    XlflowAssert.AssertTrue result.ContentLengthKnown
    XlflowAssert.AssertEquals CCur(1048576), result.ContentLength
    XlflowAssert.AssertTrue sink.CallCount > 0
    XlflowAssert.AssertEquals CCur(1048576), sink.LastBytesTransferred
    XlflowAssert.AssertEquals CCur(1048576), sink.LastTotalBytes
    XlflowAssert.AssertEquals 1048576, FileLen(destination)
    XlflowAssert.AssertEquals 0, ReadDownloadByte(destination, 0)
    XlflowAssert.AssertEquals 250, ReadDownloadByte(destination, 250)
    DeleteDownloadFile destination
    Exit Sub

Cleanup:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    DeleteDownloadFile destination
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

'@Tag("integration")
Public Sub Test_NativeTransport_StreamsUnknownLengthAndReplacesAtomically()
    Dim client As New HttpClient
    Dim options As New HttpExecutionOptions
    Dim sink As New RecordingProgressSink
    Dim result As HttpDownloadResult
    Dim destination As String
    Dim temporaryCountBefore As Long
    Dim temporaryCountAfter As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    destination = NewDownloadDestination("stream")
    On Error GoTo Cleanup
    WriteDownloadSentinel destination
    temporaryCountBefore = CountDownloadTemporaryFiles(destination)
    client.BaseUrl = RequireBaseUrl()
    Set client.Transport = New WinHttpNativeTransport
    Set result = client.DownloadFile("/stream/2097152", destination, options, sink)
    temporaryCountAfter = CountDownloadTemporaryFiles(destination)

    XlflowAssert.AssertTrue result.Published
    XlflowAssert.AssertFalse result.ContentLengthKnown
    XlflowAssert.AssertEquals - 1, result.ContentLength
    XlflowAssert.AssertEquals CCur(2097152), result.BytesWritten
    XlflowAssert.AssertEquals - 1, sink.LastTotalBytes
    XlflowAssert.AssertTrue sink.CallCount > 1
    XlflowAssert.AssertEquals 2097152, FileLen(destination)
    XlflowAssert.AssertEquals 0, ReadDownloadByte(destination, 0)
    XlflowAssert.AssertEquals temporaryCountBefore, temporaryCountAfter
    DeleteDownloadFile destination
    Exit Sub

Cleanup:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    DeleteDownloadFile destination
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

'@Tag("integration")
Public Sub Test_NativeTransport_CancellationPreservesDestinationAndCleansTemp()
    Dim client As New HttpClient
    Dim options As New HttpExecutionOptions
    Dim token As New HttpCancellationToken
    Dim sink As New RecordingProgressSink
    Dim destination As String
    Dim temporaryCountBefore As Long
    Dim temporaryCountAfter As Long
    Dim observedNumber As Long
    Dim observedSource As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    destination = NewDownloadDestination("cancel")
    On Error GoTo Cleanup
    WriteDownloadSentinel destination
    temporaryCountBefore = CountDownloadTemporaryFiles(destination)
    sink.CancelAfter = 65536
    Set sink.CancellationToken = token
    Set options.CancellationToken = token
    client.BaseUrl = RequireBaseUrl()
    Set client.Transport = New WinHttpNativeTransport

    observedNumber = CaptureDownloadError(client, "/stream/104857600", destination, options, sink, observedSource)
    XlflowAssert.AssertEquals HttpErrCancelled, observedNumber
    XlflowAssert.AssertEquals "WinHttpNativeTransport.DownloadFile", observedSource
    temporaryCountAfter = CountDownloadTemporaryFiles(destination)
    XlflowAssert.AssertEquals temporaryCountBefore, temporaryCountAfter
    XlflowAssert.AssertEquals 8, FileLen(destination)
    XlflowAssert.AssertEquals 239, ReadDownloadByte(destination, 0)
    DeleteDownloadFile destination
    Exit Sub

Cleanup:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    DeleteDownloadFile destination
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

'@Tag("integration")
Public Sub Test_NativeTransport_NonSuccessDoesNotPublishDestination()
    Dim client As New HttpClient
    Dim options As New HttpExecutionOptions
    Dim sink As New RecordingProgressSink
    Dim result As HttpDownloadResult
    Dim destination As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    destination = NewDownloadDestination("status")
    On Error GoTo Cleanup
    WriteDownloadSentinel destination
    client.BaseUrl = RequireBaseUrl()
    Set client.Transport = New WinHttpNativeTransport
    Set result = client.DownloadFile("/status/503", destination, options, sink)

    XlflowAssert.AssertEquals 503, result.StatusCode
    XlflowAssert.AssertFalse result.IsSuccess
    XlflowAssert.AssertFalse result.Published
    XlflowAssert.AssertEquals 0, result.BytesWritten
    XlflowAssert.AssertEquals 0, sink.CallCount
    XlflowAssert.AssertEquals 8, FileLen(destination)
    XlflowAssert.AssertEquals 239, ReadDownloadByte(destination, 0)
    DeleteDownloadFile destination
    Exit Sub

Cleanup:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    DeleteDownloadFile destination
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub AssertTextEcho(ByVal Method As String)
    Dim client As New HttpClient
    Dim Request As New HttpRequest
    Dim Body As New HttpBody
    Dim response As HttpResponse

    Body.Text = "native-payload-" & LCase$(Method)
    Request.Method = Method
    Request.Url = RequireBaseUrl() & "/echo"
    Set Request.Body = Body
    Set client.Transport = New WinHttpNativeTransport

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

Private Function CaptureDownloadError(ByVal client As HttpClient, ByVal Url As String, ByVal DestinationPath As String, ByVal Options As HttpExecutionOptions, ByVal Progress As IHttpProgressSink, ByRef errorSource As String) As Long
    Dim result As HttpDownloadResult

    On Error GoTo Failed
    Set result = client.DownloadFile(Url, DestinationPath, Options, Progress)
    If result Is Nothing Then Err.Raise HttpErrProtocol, "CaptureDownloadError", "Download returned Nothing."
    Exit Function

Failed:
    CaptureDownloadError = Err.Number
    errorSource = Err.Source
    Err.Clear
End Function
