Attribute VB_Name = "WinHttpCancellationStressTests"
Option Explicit

Private Const ComCancellationRequests As Long = 4
Private Const ComCancellationDelayMilliseconds As Long = 2000
Private Const ComDeadlineRequests As Long = 4
Private Const ComDeadlineMilliseconds As Long = 25
Private Const ComDeadlineDelayMilliseconds As Long = 250
Private Const ComReceiveTimeoutMilliseconds As Long = 1000
Private Const ComReceiveTimeoutDelayMilliseconds As Long = 10000
Private Const NativeDownloadBytes As Long = 65536
Private Const DownloadCancelAfterBytes As Currency = 65536

Private mScheduledCancellation As HttpCancellationToken
Private mScheduledCancellationTime As Date

'@Tag("stress")
Public Sub Test_CancellationStress_ComActiveCancellation()
    Dim client As New HttpClient
    Dim baseUrl As String
    Dim startMarker As String
    Dim doneMarker As String
    Dim releaseMarker As String
    Dim iterations As Long
    Dim index As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If Not ReadScenarioEnvironment(baseUrl, startMarker, doneMarker, releaseMarker, iterations) Then Exit Sub
    On Error GoTo Cleanup
    client.BaseUrl = baseUrl
    WriteRecoveryRequest client
    RunComActiveCancellation client, baseUrl
    WriteRecoveryRequest client
    WriteMarker startMarker

    For index = 1 To iterations
        RunComActiveCancellation client, baseUrl
        WriteRecoveryRequest client
    Next index

    WriteMarker doneMarker
    WaitForRelease releaseMarker, "WinHttpCancellationStressTests.ComActiveCancellation"
    Exit Sub

    Cleanup: ' xlflow:disable-line VBA237
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    ClearScheduledCancellation
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorDescription
End Sub

'@Tag("stress")
Public Sub Test_CancellationStress_ComDeadline()
    Dim client As New HttpClient
    Dim baseUrl As String
    Dim startMarker As String
    Dim doneMarker As String
    Dim releaseMarker As String
    Dim iterations As Long
    Dim index As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If Not ReadScenarioEnvironment(baseUrl, startMarker, doneMarker, releaseMarker, iterations) Then Exit Sub
    On Error GoTo Cleanup
    client.BaseUrl = baseUrl
    WriteRecoveryRequest client
    RunComDeadline client, baseUrl
    WriteRecoveryRequest client
    WriteMarker startMarker

    For index = 1 To iterations
        RunComDeadline client, baseUrl
        WriteRecoveryRequest client
    Next index

    WriteMarker doneMarker
    WaitForRelease releaseMarker, "WinHttpCancellationStressTests.ComDeadline"
    Exit Sub

    Cleanup: ' xlflow:disable-line VBA237
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    ClearScheduledCancellation
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorDescription
End Sub

'@Tag("stress")
Public Sub Test_CancellationStress_NativeDownloadCancellation()
    Dim client As New HttpClient
    Dim baseUrl As String
    Dim startMarker As String
    Dim doneMarker As String
    Dim releaseMarker As String
    Dim iterations As Long
    Dim index As Long
    Dim destination As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If Not ReadScenarioEnvironment(baseUrl, startMarker, doneMarker, releaseMarker, iterations) Then Exit Sub
    On Error GoTo Cleanup
    client.BaseUrl = baseUrl
    Set client.Transport = New WinHttpNativeTransport
    WriteRecoveryRequest client
    RunNativeDownloadCancellation client, "warmup"
    WriteRecoveryRequest client
    WriteMarker startMarker

    For index = 1 To iterations
        RunNativeDownloadCancellation client, CStr(index)
        WriteRecoveryRequest client
    Next index

    WriteMarker doneMarker
    WaitForRelease releaseMarker, "WinHttpCancellationStressTests.NativeDownloadCancellation"
    Exit Sub

    Cleanup: ' xlflow:disable-line VBA237
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    DeleteDownloadFile destination
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorDescription
End Sub

'@Tag("stress")
Public Sub Test_CancellationStress_ComReceiveTimeout()
    Dim client As New HttpClient
    Dim baseUrl As String
    Dim startMarker As String
    Dim doneMarker As String
    Dim releaseMarker As String
    Dim iterations As Long
    Dim index As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If Not ReadScenarioEnvironment(baseUrl, startMarker, doneMarker, releaseMarker, iterations) Then Exit Sub
    On Error GoTo Cleanup
    client.BaseUrl = baseUrl
    WriteRecoveryRequest client
    RunComReceiveTimeout client, baseUrl
    WriteRecoveryRequest client
    WriteMarker startMarker

    For index = 1 To iterations
        RunComReceiveTimeout client, baseUrl
        WriteRecoveryRequest client
    Next index

    WriteMarker doneMarker
    WaitForRelease releaseMarker, "WinHttpCancellationStressTests.ComReceiveTimeout"
    Exit Sub

    Cleanup: ' xlflow:disable-line VBA237
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorDescription
End Sub

Public Sub CancelScheduledRequest()
    If Not mScheduledCancellation Is Nothing Then mScheduledCancellation.Cancel
End Sub

Private Function ReadScenarioEnvironment(ByRef baseUrl As String, ByRef startMarker As String, ByRef doneMarker As String, ByRef releaseMarker As String, ByRef iterations As Long) As Boolean
    baseUrl = Trim$(Environ$("VBA_HTTP_TEST_BASE_URL"))
    startMarker = Trim$(Environ$("VBA_HTTP_CANCELLATION_START_PATH"))
    doneMarker = Trim$(Environ$("VBA_HTTP_CANCELLATION_DONE_PATH"))
    releaseMarker = Trim$(Environ$("VBA_HTTP_CANCELLATION_RELEASE_PATH"))
    If Len(baseUrl) = 0 Or Len(startMarker) = 0 Or Len(doneMarker) = 0 Or Len(releaseMarker) = 0 Then
        XlflowAssert.AssertInconclusive "Cancellation stress environment is incomplete; run task test:cancellation-stress."
        Exit Function
    End If
    iterations = ReadBoundedLong("VBA_HTTP_CANCELLATION_ITERATIONS", 25, 1, 1000)
    ReadScenarioEnvironment = True
End Function

Private Function CancellationRequests(ByVal baseUrl As String, ByVal count As Long, ByVal delayMilliseconds As Long) As Collection
    Set CancellationRequests = BuildRequests(baseUrl, count, "/delay/" & CStr(delayMilliseconds))
End Function

Private Function DeadlineRequests(ByVal baseUrl As String, ByVal count As Long, ByVal delayMilliseconds As Long) As Collection
    Set DeadlineRequests = BuildRequests(baseUrl, count, "/delay/" & CStr(delayMilliseconds))
End Function

Private Function BuildRequests(ByVal baseUrl As String, ByVal count As Long, ByVal resource As String) As Collection
    Dim Requests As New Collection
    Dim Request As HttpRequest
    Dim index As Long

    For index = 1 To count
        Set Request = New HttpRequest
        Request.Method = "GET"
        Request.Url = baseUrl & resource
        Requests.Add Request
    Next index
    Set BuildRequests = Requests
End Function

Private Sub RunComActiveCancellation(ByVal client As HttpClient, ByVal baseUrl As String)
    Dim Requests As Collection
    Dim Options As New HttpBatchOptions
    Dim policy As New HttpRetryPolicy
    Dim token As New HttpCancellationToken
    Dim result As HttpBatchResult

    Set Requests = CancellationRequests(baseUrl, ComCancellationRequests, ComCancellationDelayMilliseconds)
    Options.MaxConcurrency = ComCancellationRequests
    Options.PollIntervalMilliseconds = 5
    Options.YieldToHost = True
    Options.YieldIntervalMilliseconds = 20
    policy.MaxAttempts = 1
    Set Options.RetryPolicy = policy
    Set Options.CancellationToken = token
    Set mScheduledCancellation = token
    mScheduledCancellationTime = DateAdd("s", 1, Now)
    Application.OnTime mScheduledCancellationTime, "WinHttpCancellationStressTests.CancelScheduledRequest"
    Set result = client.ExecuteMany(Requests, Options)
    ClearScheduledCancellation
    XlflowAssert.AssertEquals ComCancellationRequests, result.Count
    XlflowAssert.AssertEquals ComCancellationRequests, result.CancelledCount
    XlflowAssert.AssertEquals 0, result.FailureCount
End Sub

Private Sub RunComDeadline(ByVal client As HttpClient, ByVal baseUrl As String)
    Dim Requests As Collection
    Dim Options As New HttpBatchOptions
    Dim policy As New HttpRetryPolicy
    Dim result As HttpBatchResult
    Dim itemIndex As Long

    Set Requests = DeadlineRequests(baseUrl, ComDeadlineRequests, ComDeadlineDelayMilliseconds)
    Options.MaxConcurrency = ComDeadlineRequests
    Options.PollIntervalMilliseconds = 5
    Options.RequestDeadlineMilliseconds = ComDeadlineMilliseconds
    Options.YieldToHost = False
    policy.MaxAttempts = 1
    Set Options.RetryPolicy = policy
    Set result = client.ExecuteMany(Requests, Options)
    XlflowAssert.AssertEquals ComDeadlineRequests, result.Count
    XlflowAssert.AssertEquals ComDeadlineRequests, result.FailureCount
    XlflowAssert.AssertEquals 0, result.CancelledCount
    For itemIndex = 1 To result.Count
        XlflowAssert.AssertEquals HttpErrorTimeout, result.ItemAt(itemIndex).ErrorCategory
    Next itemIndex
End Sub

Private Sub RunComReceiveTimeout(ByVal client As HttpClient, ByVal baseUrl As String)
    Dim Request As New HttpRequest
    Dim observedNumber As Long

    Request.Method = "GET"
    Request.Url = baseUrl & "/delay/" & CStr(ComReceiveTimeoutDelayMilliseconds)
    Request.Timeouts.ResolveMilliseconds = ComReceiveTimeoutMilliseconds
    Request.Timeouts.ConnectMilliseconds = ComReceiveTimeoutMilliseconds
    Request.Timeouts.SendMilliseconds = ComReceiveTimeoutMilliseconds
    Request.Timeouts.ReceiveMilliseconds = ComReceiveTimeoutMilliseconds
    observedNumber = CaptureRequestFailure(client, Request)
    XlflowAssert.AssertEquals HttpErrorTimeout, HttpErrors.CategoryFromNumber(observedNumber)
End Sub

Private Sub RunNativeDownloadCancellation(ByVal client As HttpClient, ByVal suffix As String)
    Dim destination As String
    Dim Options As New HttpExecutionOptions
    Dim token As New HttpCancellationToken
    Dim sink As New RecordingProgressSink
    Dim observedNumber As Long
    Dim temporaryCountBefore As Long
    Dim temporaryCountAfter As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    destination = NewDownloadDestination("cancel-stress-" & suffix)
    On Error GoTo Cleanup
    WriteDownloadSentinel destination
    temporaryCountBefore = CountDownloadTemporaryFiles(destination)
    sink.CancelAfter = DownloadCancelAfterBytes
    Set sink.CancellationToken = token
    Set Options.CancellationToken = token
    observedNumber = CaptureDownloadFailure(client, "/stream/" & CStr(NativeDownloadBytes), destination, Options, sink)
    XlflowAssert.AssertEquals HttpErrCancelled, observedNumber
    temporaryCountAfter = CountDownloadTemporaryFiles(destination)
    XlflowAssert.AssertEquals temporaryCountBefore, temporaryCountAfter
    XlflowAssert.AssertEquals 8, FileLen(destination)
    XlflowAssert.AssertEquals 239, ReadDownloadByte(destination, 0)
    XlflowAssert.AssertEquals 232, ReadDownloadByte(destination, 7)
    DeleteDownloadFile destination
    Exit Sub

    Cleanup: ' xlflow:disable-line VBA237
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    DeleteDownloadFile destination
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub WriteRecoveryRequest(ByVal client As HttpClient)
    Dim response As HttpResponse

    Set response = client.GetResponse("/status/204")
    XlflowAssert.AssertEquals 204, response.StatusCode
End Sub

Private Function CaptureDownloadFailure(ByVal client As HttpClient, ByVal Url As String, ByVal destination As String, ByVal Options As HttpExecutionOptions, ByVal Progress As IHttpProgressSink) As Long
    On Error GoTo Failed
    Call client.DownloadFile(Url, destination, Options, Progress)
    CaptureDownloadFailure = 0
    Exit Function

Failed:
    CaptureDownloadFailure = Err.Number
    Err.Clear
End Function

Private Function CaptureRequestFailure(ByVal client As HttpClient, ByVal Request As HttpRequest) As Long
    On Error GoTo Failed
    Call client.Execute(Request)
    CaptureRequestFailure = 0
    Exit Function

Failed:
    CaptureRequestFailure = Err.Number
    Err.Clear
End Function

Private Sub ClearScheduledCancellation()
    If mScheduledCancellation Is Nothing Then Exit Sub
    On Error Resume Next ' xlflow:disable-line VBA214
    Application.OnTime mScheduledCancellationTime, "WinHttpCancellationStressTests.CancelScheduledRequest", , False
    On Error GoTo 0
    Set mScheduledCancellation = Nothing
End Sub

Private Function ReadBoundedLong(ByVal Name As String, ByVal defaultValue As Long, ByVal minimum As Long, ByVal maximum As Long) As Long
    Dim raw As String

    raw = Trim$(Environ$(Name))
    If Len(raw) = 0 Then
        ReadBoundedLong = defaultValue
        Exit Function
    End If
    If Not IsNumeric(raw) Then HttpErrors.RaiseValidation "WinHttpCancellationStressTests", Name & " must be numeric."
    On Error GoTo InvalidValue
    ReadBoundedLong = CLng(raw)
    On Error GoTo 0
    If ReadBoundedLong < minimum Or ReadBoundedLong > maximum Then HttpErrors.RaiseValidation "WinHttpCancellationStressTests", Name & " is outside the supported range."
    Exit Function

InvalidValue:
    On Error GoTo 0
    HttpErrors.RaiseValidation "WinHttpCancellationStressTests", Name & " is not a valid 32-bit integer."
End Function

Private Sub WriteMarker(ByVal path As String)
    Dim fileSystem As Object
    Dim stream As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem Is Nothing Then Err.Raise HttpErrIo, "WinHttpCancellationStressTests.WriteMarker", "File-system marker support is unavailable."
    Set stream = CallByName(fileSystem, "CreateTextFile", VbMethod, path, True, False)
    If stream Is Nothing Then Err.Raise HttpErrIo, "WinHttpCancellationStressTests.WriteMarker", "Could not create the cancellation stress marker."
    CallByName stream, "WriteLine", VbMethod, "ready"
    CallByName stream, "Close", VbMethod
    Set stream = Nothing
    Set fileSystem = Nothing
End Sub

Private Sub WaitForRelease(ByVal path As String, ByVal Source As String)
    Dim fileSystem As Object
    Dim deadline As Date

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem Is Nothing Then Err.Raise HttpErrIo, Source, "File-system release gate support is unavailable."
    deadline = DateAdd("s", 180, Now)
    Do While Not CBool(CallByName(fileSystem, "FileExists", VbMethod, path))
        If Now >= deadline Then Err.Raise HttpErrTimeout, Source, "Cancellation stress release gate was not opened."
        HttpTiming.Pause 50
    Loop
    Set fileSystem = Nothing
End Sub
