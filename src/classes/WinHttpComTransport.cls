VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpComTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHttpTransport
Implements IHttpBatchTransport
Implements IHttpAttemptTransport

Private Const TransportSource As String = "WinHttpComTransport.Execute"
Private Const WinHttpOptionEnableRedirects As Long = 6
Private Const WinHttpOptionEnableHttpsToHttpRedirects As Long = 12
Private Const WinHttpOptionMaxAutomaticRedirects As Long = 14
Private Const WinHttpRequestProxySettingDefault As Long = 0
Private Const WinHttpRequestProxySettingDirect As Long = 1
Private Const WinHttpRequestProxySettingProxy As Long = 2
Private Const ComAbortDrainMilliseconds As Long = 250
Private Const ComAbortCancellationDrainMilliseconds As Long = 250

Public Function Execute(ByVal Request As HttpRequest) As HttpResponse
    Set Execute = ExecuteCore(Request)
End Function

Private Function IHttpTransport_Execute(ByVal Request As HttpRequest) As HttpResponse
    Set IHttpTransport_Execute = ExecuteCore(Request)
End Function

Private Function IHttpBatchTransport_ExecuteMany(ByVal Requests As Collection, ByVal Options As HttpBatchOptions) As HttpBatchResult
    Set IHttpBatchTransport_ExecuteMany = ExecuteManyCore(Requests, Options)
End Function

Private Function IHttpAttemptTransport_ExecuteAttempt(ByVal Request As HttpRequest, ByVal Options As HttpAttemptOptions) As HttpResponse
    Set IHttpAttemptTransport_ExecuteAttempt = ExecuteAttemptCore(Request, Options)
End Function

Private Function ExecuteCore(ByVal Request As HttpRequest) As HttpResponse
    Dim backend As Object
    Dim started As Currency
    Dim finished As Currency
    Dim response As HttpResponse
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo TransportFailure
    If Request Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Request cannot be Nothing."
    End If
    ValidateProtocolOptions Request
    ValidateDecompressionOptions Request

    Set backend = CreateObject("WinHttp.WinHttpRequest.5.1")
    backend.SetTimeouts Request.Timeouts.ResolveMilliseconds, Request.Timeouts.ConnectMilliseconds, Request.Timeouts.SendMilliseconds, Request.Timeouts.ReceiveMilliseconds
    backend.Option(WinHttpOptionEnableRedirects) = Request.FollowRedirects And Not HttpSecurity.HasRedirectSensitiveHeaders(Request.Headers)
    backend.Option(WinHttpOptionEnableHttpsToHttpRedirects) = False
    backend.Option(WinHttpOptionMaxAutomaticRedirects) = Request.MaxRedirects

    started = HttpTiming.CounterValue()
    backend.Open Request.Method, Request.Url, False
    ApplyProxy backend, Request.ProxyOptions, TransportSource
    ApplyChallengeCredentials backend, Request, Request.AuthProvider, TransportSource
    ApplyRequestHeaders backend, Request.Headers
    SendRequest backend, Request.Body
    finished = HttpTiming.CounterValue()

    Set response = BuildResponse(backend, HttpTiming.ElapsedMilliseconds(started, finished))
    Set backend = Nothing
    Set ExecuteCore = response
    Exit Function

TransportFailure:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    If Not TryReleaseFailedBackend(backend) Then
        errorDescription = errorDescription & " WinHTTP abort cleanup also failed."
    End If

    If HttpErrors.CategoryFromNumber(errorNumber) <> HttpErrorNone Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If
    RaiseMappedFailure errorNumber
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Function ExecuteAttemptCore(ByVal Request As HttpRequest, ByVal Options As HttpAttemptOptions) As HttpResponse
    Dim operation As WinHttpComOperation
    Dim response As HttpResponse
    Dim lastYield As Currency
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim abortNumber As Long
    Dim abortSource As String
    Dim abortDescription As String

    On Error GoTo AttemptFailed
    If Request Is Nothing Then HttpErrors.RaiseValidation TransportSource, "Request cannot be Nothing."
    If Options Is Nothing Then HttpErrors.RaiseValidation TransportSource, "Attempt options cannot be Nothing."
    ValidateProtocolOptions Request
    ValidateDecompressionOptions Request
    If AttemptCancellationRequested(Options) Then Err.Raise HttpErrCancelled, TransportSource, "HTTP operation was cancelled."

    Set operation = CreateAsyncOperation(Request, 1, Options.DeadlineMilliseconds)
    lastYield = HttpTiming.CounterValue()
    Do
        If AttemptCancellationRequested(Options) Then
            If Not TryAbortOperation(operation, abortNumber, abortSource, abortDescription, ComAbortCancellationDrainMilliseconds) Then Err.Raise abortNumber, abortSource, abortDescription
            Err.Raise HttpErrCancelled, TransportSource, "HTTP operation was cancelled."
        End If
        If operation.DeadlineExceeded(HttpTiming.CounterValue()) Then
            If Not TryAbortOperation(operation, abortNumber, abortSource, abortDescription) Then Err.Raise abortNumber, abortSource, abortDescription
            Err.Raise HttpErrTimeout, TransportSource, "HTTP attempt deadline expired."
        End If
        If PollAsyncOperation(operation) Then
            Set response = FinishAsyncOperation(operation)
            Set ExecuteAttemptCore = response
            Exit Function
        End If
        If Options.YieldToHost And HttpTiming.ElapsedMilliseconds(lastYield, HttpTiming.CounterValue()) >= Options.YieldIntervalMilliseconds Then
            ' Cooperative host yielding is an explicit ADR-0005 attempt boundary.
            DoEvents ' xlflow:disable-line VB007
            lastYield = HttpTiming.CounterValue()
        End If
        HttpTiming.Pause Options.PollIntervalMilliseconds
    Loop

AttemptFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    If Not operation Is Nothing Then
        If Not operation.Backend Is Nothing Then
            If Not TryAbortOperation(operation, abortNumber, abortSource, abortDescription, ComAbortDrainMilliseconds) Then
                errorDescription = errorDescription & " WinHTTP abort cleanup also failed."
            End If
        End If
    End If
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Function AttemptCancellationRequested(ByVal Options As HttpAttemptOptions) As Boolean
    If Options.CancellationToken Is Nothing Then Exit Function
    AttemptCancellationRequested = Options.CancellationToken.IsCancellationRequested
End Function

Private Function ExecuteManyCore(ByVal Requests As Collection, ByVal Options As HttpBatchOptions) As HttpBatchResult
    Dim result As New HttpBatchResult
    Dim running As New Collection
    Dim workItems() As HttpBatchWorkItem
    Dim work As HttpBatchWorkItem
    Dim operation As WinHttpComOperation
    Dim response As HttpResponse
    Dim index As Long
    Dim completed As Boolean
    Dim completedCount As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim abortNumber As Long
    Dim abortSource As String
    Dim abortDescription As String
    Dim abortSucceeded As Boolean
    Dim now As Currency
    Dim lastYield As Currency
    Dim totalStarted As Currency
    Dim elapsed As Double
    Dim attemptDeadline As Long
    Dim policy As HttpRetryPolicy
    Dim runtime As IHttpReliabilityRuntime

    If Requests Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Requests cannot be Nothing."
        Exit Function
    End If
    If Options Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Batch options cannot be Nothing."
        Exit Function
    End If
    If Options.RetryPolicy Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Batch retry policy cannot be Nothing."
        Exit Function
    End If
    If Requests.Count = 0 Then
        Set ExecuteManyCore = result
        Exit Function
    End If

    Set policy = Options.RetryPolicy.Clone()
    Set runtime = New HttpSystemReliabilityRuntime
    ReDim workItems(1 To Requests.Count)
    For index = 1 To Requests.Count
        Set work = New HttpBatchWorkItem
        work.Initialize index, Requests(index)
        Set workItems(index) = work
    Next index
    totalStarted = HttpTiming.CounterValue()
    lastYield = totalStarted

    Do While completedCount < Requests.Count
        If CancellationRequested(Options) Then
            CompleteOutstanding running, workItems, True, "HTTP batch operation was cancelled."
            Exit Do
        End If

        now = HttpTiming.CounterValue()
        elapsed = HttpTiming.ElapsedMilliseconds(totalStarted, now)
        If Options.TotalDeadlineMilliseconds > 0 And elapsed >= Options.TotalDeadlineMilliseconds Then
            CompleteOutstanding running, workItems, False, "HTTP batch total deadline expired."
            Exit Do
        End If

        For index = running.Count To 1 Step - 1
            Set operation = running(index)
            Set work = workItems(operation.Index)
            completed = False
            If operation.DeadlineExceeded(now) Then
                abortSucceeded = TryAbortOperation(operation, abortNumber, abortSource, abortDescription)
                running.Remove index
                If abortSucceeded Then
                    HandleBatchFailure work, policy, HttpErrTimeout, TransportSource, "HTTP batch request deadline exceeded.", elapsed, runtime, completedCount
                Else
                    HandleBatchFailure work, policy, abortNumber, abortSource, abortDescription, elapsed, runtime, completedCount
                End If
            ElseIf Not TryPollOperation(operation, completed, errorNumber, errorSource, errorDescription) Then
                abortSucceeded = TryAbortOperation(operation, abortNumber, abortSource, abortDescription, ComAbortDrainMilliseconds)
                If Not abortSucceeded Then errorDescription = errorDescription & " WinHTTP abort cleanup also failed."
                running.Remove index
                HandleBatchFailure work, policy, errorNumber, errorSource, errorDescription, elapsed, runtime, completedCount
            ElseIf completed Then
                running.Remove index
                If TryFinishOperation(operation, response, errorNumber, errorSource, errorDescription) Then
                    HandleBatchResponse work, policy, response, elapsed, runtime, completedCount
                Else
                    HandleBatchFailure work, policy, errorNumber, errorSource, errorDescription, elapsed, runtime, completedCount
                End If
            End If
        Next index

        For index = 1 To Requests.Count
            If running.Count >= Options.MaxConcurrency Then Exit For
            Set work = workItems(index)
            If work.IsEligible(elapsed) Then
                work.BeginAttempt
                attemptDeadline = EffectiveAttemptDeadline(Options, elapsed)
                Set operation = Nothing
                If TryStartOperation(work.RequestSnapshot, work.Index, attemptDeadline, operation, errorNumber, errorSource, errorDescription) Then
                    running.Add operation
                Else
                    HandleBatchFailure work, policy, errorNumber, errorSource, errorDescription, elapsed, runtime, completedCount
                End If
            End If
        Next index

        If completedCount < Requests.Count Then
            If Options.YieldToHost And HttpTiming.ElapsedMilliseconds(lastYield, HttpTiming.CounterValue()) >= Options.YieldIntervalMilliseconds Then
                ' Cooperative host yielding is an explicit ADR-0004 scheduler boundary.
                DoEvents ' xlflow:disable-line VB007
                lastYield = HttpTiming.CounterValue()
            End If
            HttpTiming.Pause Options.PollIntervalMilliseconds
        End If
    Loop

    For index = 1 To Requests.Count
        Set work = workItems(index)
        If work.TerminalItem Is Nothing Then Err.Raise HttpErrProtocol, TransportSource, "Batch scheduler left an item incomplete."
        result.Add work.TerminalItem
    Next index
    Set ExecuteManyCore = result
End Function

Private Sub HandleBatchResponse(ByVal work As HttpBatchWorkItem, ByVal policy As HttpRetryPolicy, ByVal response As HttpResponse, ByVal elapsed As Double, ByVal runtime As IHttpReliabilityRuntime, ByRef completedCount As Long)
    Dim delayMilliseconds As Long

    If policy.AllowsMethod(work.RequestSnapshot.Method) And policy.AllowsStatus(response.StatusCode) And work.Attempt < policy.MaxAttempts Then
        delayMilliseconds = BatchResponseDelay(policy, response, work.Attempt, runtime)
        work.ScheduleRetry elapsed, delayMilliseconds
    Else
        work.CompleteSuccess response
        completedCount = completedCount + 1
    End If
End Sub

Private Sub HandleBatchFailure(ByVal work As HttpBatchWorkItem, ByVal policy As HttpRetryPolicy, ByVal errorNumber As Long, ByVal errorSource As String, ByVal errorDescription As String, ByVal elapsed As Double, ByVal runtime As IHttpReliabilityRuntime, ByRef completedCount As Long)
    Dim category As HttpErrorCategory
    Dim delayMilliseconds As Long

    category = HttpErrors.CategoryFromNumber(errorNumber)
    If policy.AllowsMethod(work.RequestSnapshot.Method) And policy.AllowsError(category) And work.Attempt < policy.MaxAttempts Then
        delayMilliseconds = policy.BackoffMilliseconds(work.Attempt, runtime.NextJitterUnit())
        work.ScheduleRetry elapsed, delayMilliseconds
    Else
        work.CompleteFailure errorNumber, errorSource, errorDescription
        completedCount = completedCount + 1
    End If
End Sub

Private Function BatchResponseDelay(ByVal policy As HttpRetryPolicy, ByVal response As HttpResponse, ByVal completedAttempt As Long, ByVal runtime As IHttpReliabilityRuntime) As Long
    Dim serverDelay As Long

    If policy.RespectRetryAfter Then
        If HttpRetryAfter.TryParse(response.Headers.GetValue("Retry-After"), runtime.UtcNow(), serverDelay) Then
            If serverDelay > policy.MaxDelayMilliseconds Then serverDelay = policy.MaxDelayMilliseconds
            BatchResponseDelay = serverDelay
            Exit Function
        End If
    End If
    BatchResponseDelay = policy.BackoffMilliseconds(completedAttempt, runtime.NextJitterUnit())
End Function

Private Function EffectiveAttemptDeadline(ByVal Options As HttpBatchOptions, ByVal elapsed As Double) As Long
    Dim remaining As Double

    If Options.TotalDeadlineMilliseconds > 0 Then
        remaining = CDbl(Options.TotalDeadlineMilliseconds) - elapsed
        If remaining <= 0 Then Exit Function
        If remaining > 2147483647# Then
            EffectiveAttemptDeadline = 2147483647
        Else
            EffectiveAttemptDeadline = CLng(Fix(remaining + 0.999999))
        End If
    End If
    If Options.RequestDeadlineMilliseconds > 0 Then
        If EffectiveAttemptDeadline = 0 Or Options.RequestDeadlineMilliseconds < EffectiveAttemptDeadline Then
            EffectiveAttemptDeadline = Options.RequestDeadlineMilliseconds
        End If
    End If
End Function

Private Sub CompleteOutstanding(ByVal running As Collection, ByRef workItems() As HttpBatchWorkItem, ByVal cancelled As Boolean, ByVal description As String)
    Dim operation As WinHttpComOperation
    Dim work As HttpBatchWorkItem
    Dim index As Long
    Dim abortNumber As Long
    Dim abortSource As String
    Dim abortDescription As String
    Dim drainMilliseconds As Long

    If running Is Nothing Then
        Err.Raise HttpErrProtocol, TransportSource, "Batch scheduler lost its running-operation collection."
        Exit Sub
    End If
    For index = running.Count To 1 Step - 1
        Set operation = running(index)
        Set work = workItems(operation.Index)
        If cancelled Then drainMilliseconds = ComAbortCancellationDrainMilliseconds
        If TryAbortOperation(operation, abortNumber, abortSource, abortDescription, drainMilliseconds) Then
            If cancelled Then
                work.CompleteCancelled
            Else
                work.CompleteFailure HttpErrTimeout, TransportSource, description
            End If
        Else
            work.CompleteFailure abortNumber, abortSource, abortDescription
        End If
        running.Remove index
    Next index
    For index = LBound(workItems) To UBound(workItems)
        Set work = workItems(index)
        If Not work.IsComplete Then
            If cancelled Then
                work.CompleteCancelled
            Else
                work.CompleteFailure HttpErrTimeout, TransportSource, description
            End If
        End If
    Next index
End Sub

Private Function CancellationRequested(ByVal Options As HttpBatchOptions) As Boolean
    If Options.CancellationToken Is Nothing Then Exit Function
    CancellationRequested = Options.CancellationToken.IsCancellationRequested
End Function

Private Function TryStartOperation(ByVal Request As HttpRequest, ByVal itemIndex As Long, ByVal deadlineMilliseconds As Long, ByRef operation As WinHttpComOperation, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
    On Error GoTo StartFailed
    Set operation = CreateAsyncOperation(Request, itemIndex, deadlineMilliseconds)
    TryStartOperation = True
    Exit Function

StartFailed:
    CaptureCurrentError errorNumber, errorSource, errorDescription
    TryStartOperation = False
End Function

Private Function TryPollOperation(ByVal operation As WinHttpComOperation, ByRef completed As Boolean, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
    On Error GoTo PollFailed
    completed = PollAsyncOperation(operation)
    TryPollOperation = True
    Exit Function

PollFailed:
    CaptureCurrentError errorNumber, errorSource, errorDescription
    TryPollOperation = False
End Function

Private Function TryFinishOperation(ByVal operation As WinHttpComOperation, ByRef response As HttpResponse, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
    On Error GoTo FinishFailed
    Set response = FinishAsyncOperation(operation)
    TryFinishOperation = True
    Exit Function

FinishFailed:
    CaptureCurrentError errorNumber, errorSource, errorDescription
    TryFinishOperation = False
End Function

Private Function TryAbortOperation(ByVal operation As WinHttpComOperation, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String, Optional ByVal drainMilliseconds As Long = 0) As Boolean
    On Error GoTo AbortFailed
    operation.Backend.Abort
    If drainMilliseconds > 0 Then HttpTiming.Pause drainMilliseconds
    operation.ReleaseBackend
    TryAbortOperation = True
    Exit Function

AbortFailed:
    Err.Clear
    operation.ReleaseBackend
    errorNumber = HttpErrIo
    errorSource = TransportSource
    errorDescription = "WinHTTP request abort failed."
    TryAbortOperation = False
End Function

Private Function CreateAsyncOperation(ByVal Request As HttpRequest, ByVal itemIndex As Long, ByVal deadlineMilliseconds As Long) As WinHttpComOperation
    Dim backend As Object
    Dim operation As New WinHttpComOperation
    Dim started As Currency
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo TransportFailure
    ValidateProtocolOptions Request
    ValidateDecompressionOptions Request
    Set backend = CreateObject("WinHttp.WinHttpRequest.5.1")
    backend.SetTimeouts Request.Timeouts.ResolveMilliseconds, Request.Timeouts.ConnectMilliseconds, Request.Timeouts.SendMilliseconds, Request.Timeouts.ReceiveMilliseconds
    backend.Option(WinHttpOptionEnableRedirects) = Request.FollowRedirects And Not HttpSecurity.HasRedirectSensitiveHeaders(Request.Headers)
    backend.Option(WinHttpOptionEnableHttpsToHttpRedirects) = False
    backend.Option(WinHttpOptionMaxAutomaticRedirects) = Request.MaxRedirects
    started = HttpTiming.CounterValue()
    backend.Open Request.Method, Request.Url, True
    ApplyProxy backend, Request.ProxyOptions, TransportSource
    ApplyChallengeCredentials backend, Request, Request.AuthProvider, TransportSource
    ApplyRequestHeaders backend, Request.Headers
    SendRequest backend, Request.Body
    operation.Initialize itemIndex, backend, started, deadlineMilliseconds
    Set CreateAsyncOperation = operation
    Exit Function

TransportFailure:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    If Not TryReleaseFailedBackend(backend) Then
        errorDescription = errorDescription & " WinHTTP abort cleanup also failed."
    End If
    RaiseCapturedTransportFailure errorNumber, errorSource, errorDescription
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Function TryReleaseFailedBackend(ByRef backend As Object) As Boolean
    On Error GoTo AbortFailed
    If Not backend Is Nothing Then
        backend.Abort
        ' WinHTTP completes some COM-handle teardown asynchronously after Abort.
        HttpTiming.Pause ComAbortDrainMilliseconds
    End If
    Set backend = Nothing
    TryReleaseFailedBackend = True
    Exit Function

AbortFailed:
    Err.Clear
    Set backend = Nothing
    TryReleaseFailedBackend = False
End Function

Private Sub ValidateProtocolOptions(ByVal Request As HttpRequest)
    If Request Is Nothing Then Exit Sub
    If Request.ProtocolOptions Is Nothing Then Exit Sub
    Request.ProtocolOptions.Validate
    If Request.ProtocolOptions.HasOverride Then
        HttpErrors.RaiseValidation TransportSource, "HTTP/2 and HTTP/3 protocol options require the native WinHTTP transport."
    End If
End Sub

Private Sub ValidateDecompressionOptions(ByVal Request As HttpRequest)
    If Request Is Nothing Then Exit Sub
    If Request.DecompressionOptions Is Nothing Then Exit Sub
    Request.DecompressionOptions.Validate
    If Request.DecompressionOptions.HasOverride Then
        HttpErrors.RaiseValidation TransportSource, "Response decompression options require the native WinHTTP transport."
    End If
End Sub

Private Sub ApplyProxy(ByVal backend As Object, ByVal Options As HttpProxyOptions, ByVal Source As String)
    If backend Is Nothing Then
        HttpErrors.RaiseValidation Source, "WinHTTP backend is unavailable."
        Exit Sub
    End If
    If Options Is Nothing Then Exit Sub
    Options.Validate
    Select Case Options.Mode
    Case HttpProxyDefault
        backend.SetProxy WinHttpRequestProxySettingDefault, "", ""
    Case HttpProxyNoProxy
        backend.SetProxy WinHttpRequestProxySettingDirect, "", ""
    Case HttpProxyManual
        backend.SetProxy WinHttpRequestProxySettingProxy, ProxyServerName(Options.ProxyUrl), Options.BypassList
    Case Else
        HttpErrors.RaiseValidation Source, "Unsupported proxy mode."
    End Select
End Sub

Private Sub ApplyChallengeCredentials(ByVal backend As Object, ByVal Request As HttpRequest, ByVal Provider As IHttpAuthProvider, ByVal Source As String)
    Dim challengeProvider As HttpWindowsAuthProvider

    If backend Is Nothing Or Provider Is Nothing Then Exit Sub
    If Not TypeOf Provider Is HttpWindowsAuthProvider Then Exit Sub
    Provider.Apply Request
    Set challengeProvider = Provider
    challengeProvider.ConfigureCom backend, Source
End Sub

Private Function ProxyServerName(ByVal ProxyUrl As String) As String
    Dim schemeEnd As Long

    schemeEnd = InStr(1, ProxyUrl, "://", vbBinaryCompare)
    If schemeEnd > 0 Then
        ProxyServerName = Mid$(ProxyUrl, schemeEnd + 3)
    Else
        ProxyServerName = ProxyUrl
    End If
End Function

Private Function PollAsyncOperation(ByVal operation As WinHttpComOperation) As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo TransportFailure
    PollAsyncOperation = CBool(operation.Backend.WaitForResponse(0))
    Exit Function

TransportFailure:
    CaptureCurrentError errorNumber, errorSource, errorDescription
    RaiseCapturedTransportFailure errorNumber, errorSource, errorDescription
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Function FinishAsyncOperation(ByVal operation As WinHttpComOperation) As HttpResponse
    Dim response As HttpResponse
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo TransportFailure
    Set response = BuildResponse(operation.Backend, HttpTiming.ElapsedMilliseconds(operation.Started, HttpTiming.CounterValue()))
    operation.ReleaseBackend
    Set FinishAsyncOperation = response
    Exit Function

TransportFailure:
    CaptureCurrentError errorNumber, errorSource, errorDescription
    operation.ReleaseBackend
    RaiseCapturedTransportFailure errorNumber, errorSource, errorDescription
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Sub CaptureCurrentError(ByRef number As Long, ByRef Source As String, ByRef Description As String)
    number = Err.Number
    Source = Err.Source
    Description = Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub RaiseCapturedTransportFailure(ByVal errorNumber As Long, ByVal errorSource As String, ByVal errorDescription As String)
    If HttpErrors.CategoryFromNumber(errorNumber) <> HttpErrorNone Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If
    RaiseMappedFailure errorNumber
End Sub

Private Sub ApplyRequestHeaders(ByVal backend As Object, ByVal Headers As HttpHeaders)
    Dim index As Long

    If backend Is Nothing Then
        Err.Raise HttpErrIo, TransportSource, "WinHTTP backend is unavailable."
        Exit Sub
    End If
    If Headers Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Request headers cannot be Nothing."
        Exit Sub
    End If
    For index = 1 To Headers.Count
        ' HttpHeaders validates field names and rejects CR/LF in values before this sink.
        backend.SetRequestHeader Headers.NameAt(index), Headers.ValueAt(index) ' xlflow:disable-line VBA224
    Next index
End Sub

Private Sub SendRequest(ByVal backend As Object, ByVal Body As HttpBody)
    Dim bytes As Variant

    If backend Is Nothing Then
        Err.Raise HttpErrIo, TransportSource, "WinHTTP backend is unavailable."
        Exit Sub
    End If
    If Body Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Request body cannot be Nothing."
        Exit Sub
    End If
    If Body.IsEmpty Then
        backend.Send
    Else
        bytes = Body.Bytes
        backend.Send bytes
    End If
End Sub

Private Function BuildResponse(ByVal backend As Object, ByVal elapsedMilliseconds As Double) As HttpResponse
    Dim response As New HttpResponse
    Dim Headers As HttpHeaders
    Dim Body As New HttpBody
    Dim bytes As Variant

    If backend Is Nothing Then
        Err.Raise HttpErrIo, TransportSource, "WinHTTP backend is unavailable."
        Exit Function
    End If
    Set Headers = ParseResponseHeaders(CStr(backend.GetAllResponseHeaders()))
    bytes = backend.ResponseBody
    If HttpEncoding.ByteCount(bytes) > 0 Then Body.SetBytes bytes
    response.Initialize CLng(backend.Status), CStr(backend.StatusText), Headers, Body, elapsedMilliseconds
    Set BuildResponse = response
End Function

Private Function ParseResponseHeaders(ByVal rawHeaders As String) As HttpHeaders
    Dim Headers As New HttpHeaders
    Dim lines As Variant
    Dim line As Variant
    Dim separatorIndex As Long
    Dim text As String

    lines = Split(rawHeaders, vbCrLf)
    For Each line In lines
        text = CStr(line)
        If Len(text) > 0 Then
            separatorIndex = InStr(1, text, ":", vbBinaryCompare)
            If separatorIndex <= 1 Then
                Err.Raise HttpErrProtocol, TransportSource, "WinHTTP returned a malformed response header."
            End If
            Headers.Add Left$(text, separatorIndex - 1), Trim$(Mid$(text, separatorIndex + 1))
        End If
    Next line
    Set ParseResponseHeaders = Headers
End Function

Private Sub RaiseMappedFailure(ByVal nativeErrorNumber As Long)
    Dim code As Long
    Dim category As HttpErrorCategory
    Dim summary As String

    code = nativeErrorNumber And 65535
    category = WinHttpErrorMapping.CategoryFromCode(code)
    Select Case category
    Case HttpErrorInvalidUrl
        summary = "WinHTTP rejected the request URL"
    Case HttpErrorDns
        summary = "WinHTTP could not resolve the server name"
    Case HttpErrorConnection
        summary = "WinHTTP could not establish or maintain the connection"
    Case HttpErrorTls
        summary = "WinHTTP TLS validation failed"
    Case HttpErrorTimeout
        summary = "WinHTTP request timed out"
    Case HttpErrorCancelled
        summary = "WinHTTP request was cancelled"
    Case HttpErrorProtocol
        summary = "WinHTTP reported an invalid HTTP exchange"
    Case Else
        category = HttpErrorIo
        summary = "WinHTTP request failed"
    End Select
    HttpErrors.RaiseTransport category, TransportSource, summary & " (" & CStr(code) & ")."
End Sub
