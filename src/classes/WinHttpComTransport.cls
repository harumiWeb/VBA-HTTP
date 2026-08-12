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

Private Const TransportSource As String = "WinHttpComTransport.Execute"
Private Const WinHttpOptionEnableRedirects As Long = 6
Private Const WinHttpOptionMaxAutomaticRedirects As Long = 14

Public Function Execute(ByVal Request As HttpRequest) As HttpResponse
    Set Execute = ExecuteCore(Request)
End Function

Private Function IHttpTransport_Execute(ByVal Request As HttpRequest) As HttpResponse
    Set IHttpTransport_Execute = ExecuteCore(Request)
End Function

Private Function IHttpBatchTransport_ExecuteMany(ByVal Requests As Collection, ByVal Options As HttpBatchOptions) As HttpBatchResult
    Set IHttpBatchTransport_ExecuteMany = ExecuteManyCore(Requests, Options)
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

    Set backend = CreateObject("WinHttp.WinHttpRequest.5.1")
    backend.SetTimeouts Request.Timeouts.ResolveMilliseconds, Request.Timeouts.ConnectMilliseconds, Request.Timeouts.SendMilliseconds, Request.Timeouts.ReceiveMilliseconds
    backend.Option(WinHttpOptionEnableRedirects) = Request.FollowRedirects
    backend.Option(WinHttpOptionMaxAutomaticRedirects) = Request.MaxRedirects

    started = HttpTiming.CounterValue()
    backend.Open Request.Method, Request.Url, False
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
    Set backend = Nothing

    If HttpErrors.CategoryFromNumber(errorNumber) <> HttpErrorNone Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If
    RaiseMappedFailure errorNumber
End Function

Private Function ExecuteManyCore(ByVal Requests As Collection, ByVal Options As HttpBatchOptions) As HttpBatchResult
    Dim result As New HttpBatchResult
    Dim running As New Collection
    Dim items() As HttpBatchItem
    Dim operation As WinHttpComOperation
    Dim response As HttpResponse
    Dim item As HttpBatchItem
    Dim nextIndex As Long
    Dim index As Long
    Dim completed As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim abortNumber As Long
    Dim abortSource As String
    Dim abortDescription As String
    Dim abortSucceeded As Boolean
    Dim now As Currency
    Dim lastYield As Currency

    If Requests Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Requests cannot be Nothing."
        Exit Function
    End If
    If Options Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Batch options cannot be Nothing."
        Exit Function
    End If
    If Requests.Count = 0 Then
        Set ExecuteManyCore = result
        Exit Function
    End If

    ReDim items(1 To Requests.Count)
    nextIndex = 1
    lastYield = HttpTiming.CounterValue()

    Do While nextIndex <= Requests.Count Or running.Count > 0
        If CancellationRequested(Options) Then
            CancelOutstanding running, items, nextIndex, Requests.Count
            Exit Do
        End If

        Do While nextIndex <= Requests.Count And running.Count < Options.MaxConcurrency
            Set operation = Nothing
            If TryStartOperation(Requests(nextIndex), nextIndex, Options.RequestDeadlineMilliseconds, operation, errorNumber, errorSource, errorDescription) Then
                running.Add operation
            Else
                Set item = New HttpBatchItem
                item.InitializeFailure nextIndex, errorNumber, errorSource, errorDescription
                Set items(nextIndex) = item
            End If
            nextIndex = nextIndex + 1
        Loop

        now = HttpTiming.CounterValue()
        For index = running.Count To 1 Step - 1
            Set operation = running(index)
            completed = False
            If operation.DeadlineExceeded(now) Then
                abortSucceeded = TryAbortOperation(operation, abortNumber, abortSource, abortDescription)
                Set item = New HttpBatchItem
                If abortSucceeded Then
                    item.InitializeFailure operation.Index, HttpErrTimeout, TransportSource, "HTTP batch request deadline exceeded."
                Else
                    item.InitializeFailure operation.Index, abortNumber, abortSource, abortDescription
                End If
                Set items(operation.Index) = item
                running.Remove index
            ElseIf Not TryPollOperation(operation, completed, errorNumber, errorSource, errorDescription) Then
                abortSucceeded = TryAbortOperation(operation, abortNumber, abortSource, abortDescription)
                If Not abortSucceeded Then errorDescription = errorDescription & " WinHTTP abort cleanup also failed."
                Set item = New HttpBatchItem
                item.InitializeFailure operation.Index, errorNumber, errorSource, errorDescription
                Set items(operation.Index) = item
                running.Remove index
            ElseIf completed Then
                If TryFinishOperation(operation, response, errorNumber, errorSource, errorDescription) Then
                    Set item = New HttpBatchItem
                    item.InitializeSuccess operation.Index, response
                Else
                    Set item = New HttpBatchItem
                    item.InitializeFailure operation.Index, errorNumber, errorSource, errorDescription
                End If
                Set items(operation.Index) = item
                running.Remove index
            End If
        Next index

        If running.Count > 0 Then
            If Options.YieldToHost And HttpTiming.ElapsedMilliseconds(lastYield, HttpTiming.CounterValue()) >= Options.YieldIntervalMilliseconds Then
                ' Cooperative host yielding is an explicit ADR-0004 scheduler boundary.
                DoEvents ' xlflow:disable-line VB007
                lastYield = HttpTiming.CounterValue()
            End If
            HttpTiming.Pause Options.PollIntervalMilliseconds
        End If
    Loop

    For index = 1 To Requests.Count
        If items(index) Is Nothing Then Err.Raise HttpErrProtocol, TransportSource, "Batch scheduler left an item incomplete."
        result.Add items(index)
    Next index
    Set ExecuteManyCore = result
End Function

Private Sub CancelOutstanding(ByVal running As Collection, ByRef items() As HttpBatchItem, ByRef nextIndex As Long, ByVal requestCount As Long)
    Dim operation As WinHttpComOperation
    Dim item As HttpBatchItem
    Dim index As Long
    Dim abortNumber As Long
    Dim abortSource As String
    Dim abortDescription As String

    If running Is Nothing Then
        Err.Raise HttpErrProtocol, TransportSource, "Batch scheduler lost its running-operation collection."
        Exit Sub
    End If
    For index = running.Count To 1 Step - 1
        Set operation = running(index)
        Set item = New HttpBatchItem
        If TryAbortOperation(operation, abortNumber, abortSource, abortDescription) Then
            item.InitializeCancelled operation.Index
        Else
            item.InitializeFailure operation.Index, abortNumber, abortSource, abortDescription
        End If
        Set items(operation.Index) = item
        running.Remove index
    Next index
    Do While nextIndex <= requestCount
        Set item = New HttpBatchItem
        item.InitializeCancelled nextIndex
        Set items(nextIndex) = item
        nextIndex = nextIndex + 1
    Loop
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

Private Function TryAbortOperation(ByVal operation As WinHttpComOperation, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
    On Error GoTo AbortFailed
    operation.Backend.Abort
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
    Set backend = CreateObject("WinHttp.WinHttpRequest.5.1")
    backend.SetTimeouts Request.Timeouts.ResolveMilliseconds, Request.Timeouts.ConnectMilliseconds, Request.Timeouts.SendMilliseconds, Request.Timeouts.ReceiveMilliseconds
    backend.Option(WinHttpOptionEnableRedirects) = Request.FollowRedirects
    backend.Option(WinHttpOptionMaxAutomaticRedirects) = Request.MaxRedirects
    started = HttpTiming.CounterValue()
    backend.Open Request.Method, Request.Url, True
    ApplyRequestHeaders backend, Request.Headers
    SendRequest backend, Request.Body
    operation.Initialize itemIndex, backend, started, deadlineMilliseconds
    Set CreateAsyncOperation = operation
    Exit Function

TransportFailure:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Set backend = Nothing
    RaiseCapturedTransportFailure errorNumber, errorSource, errorDescription
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
    category = CategoryFromWinHttpCode(code)
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

Private Function CategoryFromWinHttpCode(ByVal code As Long) As HttpErrorCategory
    Select Case code
    Case 12005, 12006
        CategoryFromWinHttpCode = HttpErrorInvalidUrl
    Case 12007
        CategoryFromWinHttpCode = HttpErrorDns
    Case 12029, 12030, 12031
        CategoryFromWinHttpCode = HttpErrorConnection
    Case 12037, 12038, 12044, 12045, 12055, 12057, 12157, 12175
        CategoryFromWinHttpCode = HttpErrorTls
    Case 12002
        CategoryFromWinHttpCode = HttpErrorTimeout
    Case 12017
        CategoryFromWinHttpCode = HttpErrorCancelled
    Case 12150 To 12156
        CategoryFromWinHttpCode = HttpErrorProtocol
    Case Else
        CategoryFromWinHttpCode = HttpErrorIo
    End Select
End Function
