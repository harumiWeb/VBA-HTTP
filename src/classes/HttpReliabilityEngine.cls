VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpReliabilityEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ReliabilitySource As String = "HttpClient.Execute"

Private mRuntime As IHttpReliabilityRuntime

Friend Sub Initialize(ByVal runtime As IHttpReliabilityRuntime)
    If runtime Is Nothing Then HttpErrors.RaiseValidation "HttpReliabilityEngine.Initialize", "Reliability runtime cannot be Nothing."
    Set mRuntime = runtime
End Sub

Friend Function Execute(ByVal transport As IHttpTransport, ByVal Request As HttpRequest, ByVal policy As HttpRetryPolicy, ByVal Options As HttpExecutionOptions) As HttpResponse
    Dim response As HttpResponse
    Dim started As Currency
    Dim attempt As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim category As HttpErrorCategory
    Dim delayMilliseconds As Long

    ValidateInputs transport, Request, policy, Options
    started = mRuntime.CounterValue()
    CheckControls Options, started
    ValidateControlCapability transport, Options

    For attempt = 1 To policy.MaxAttempts
        CheckControls Options, started
        If TryExecuteAttempt(transport, Request, Options, started, response, errorNumber, errorSource, errorDescription) Then
            CheckControls Options, started
            If Not policy.AllowsMethod(Request.Method) Or Not policy.AllowsStatus(response.StatusCode) Or attempt >= policy.MaxAttempts Then
                Set Execute = response
                Exit Function
            End If
            delayMilliseconds = ResponseDelay(policy, response, attempt)
        Else
            CheckControls Options, started
            category = HttpErrors.CategoryFromNumber(errorNumber)
            If Not policy.AllowsMethod(Request.Method) Or Not policy.AllowsError(category) Or attempt >= policy.MaxAttempts Then
                Err.Raise errorNumber, errorSource, errorDescription
            End If
            delayMilliseconds = policy.BackoffMilliseconds(attempt, mRuntime.NextJitterUnit())
        End If
        WaitBeforeRetry delayMilliseconds, Options, started
    Next attempt

    Err.Raise HttpErrProtocol, ReliabilitySource, "Reliability engine exhausted attempts without a terminal result."
End Function

Private Sub ValidateInputs(ByVal transport As IHttpTransport, ByVal Request As HttpRequest, ByVal policy As HttpRetryPolicy, ByVal Options As HttpExecutionOptions)
    If mRuntime Is Nothing Then HttpErrors.RaiseValidation "HttpReliabilityEngine.Execute", "Reliability runtime was not initialized."
    If transport Is Nothing Then HttpErrors.RaiseValidation "HttpReliabilityEngine.Execute", "Transport cannot be Nothing."
    If Request Is Nothing Then HttpErrors.RaiseValidation "HttpReliabilityEngine.Execute", "Request cannot be Nothing."
    If policy Is Nothing Then HttpErrors.RaiseValidation "HttpReliabilityEngine.Execute", "Retry policy cannot be Nothing."
    If Options Is Nothing Then HttpErrors.RaiseValidation "HttpReliabilityEngine.Execute", "Execution options cannot be Nothing."
    policy.Validate
End Sub

Private Sub ValidateControlCapability(ByVal transport As IHttpTransport, ByVal Options As HttpExecutionOptions)
    If (Not Options.CancellationToken Is Nothing Or Options.TotalDeadlineMilliseconds > 0) And Not TypeOf transport Is IHttpAttemptTransport Then
        HttpErrors.RaiseValidation ReliabilitySource, "Configured transport cannot enforce cancellation or total deadline during an active attempt."
    End If
End Sub

Private Function TryExecuteAttempt(ByVal transport As IHttpTransport, ByVal Request As HttpRequest, ByVal Options As HttpExecutionOptions, ByVal totalStarted As Currency, ByRef response As HttpResponse, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
    Dim attemptTransport As IHttpAttemptTransport
    Dim attemptOptions As HttpAttemptOptions

    On Error GoTo AttemptFailed
    Set response = Nothing
    If Not Options.CancellationToken Is Nothing Or Options.TotalDeadlineMilliseconds > 0 Then
        Set attemptTransport = transport
        Set attemptOptions = CreateAttemptOptions(Options, totalStarted)
        Set response = attemptTransport.ExecuteAttempt(Request, attemptOptions)
    Else
        Set response = transport.Execute(Request)
    End If
    If response Is Nothing Then Err.Raise HttpErrProtocol, ReliabilitySource, "Transport returned Nothing instead of an HttpResponse."
    TryExecuteAttempt = True
    Exit Function

AttemptFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    TryExecuteAttempt = False
End Function

Private Function CreateAttemptOptions(ByVal Options As HttpExecutionOptions, ByVal totalStarted As Currency) As HttpAttemptOptions
    Dim attemptOptions As New HttpAttemptOptions

    attemptOptions.DeadlineMilliseconds = RemainingDeadline(Options, totalStarted)
    Set attemptOptions.CancellationToken = Options.CancellationToken
    attemptOptions.PollIntervalMilliseconds = Options.PollIntervalMilliseconds
    attemptOptions.YieldToHost = Options.YieldToHost
    attemptOptions.YieldIntervalMilliseconds = Options.YieldIntervalMilliseconds
    Set CreateAttemptOptions = attemptOptions
End Function

Private Function ResponseDelay(ByVal policy As HttpRetryPolicy, ByVal response As HttpResponse, ByVal completedAttempt As Long) As Long
    Dim serverDelay As Long

    If policy.RespectRetryAfter Then
        If HttpRetryAfter.TryParse(response.Headers.GetValue("Retry-After"), mRuntime.UtcNow(), serverDelay) Then
            If serverDelay > policy.MaxDelayMilliseconds Then serverDelay = policy.MaxDelayMilliseconds
            ResponseDelay = serverDelay
            Exit Function
        End If
    End If
    ResponseDelay = policy.BackoffMilliseconds(completedAttempt, mRuntime.NextJitterUnit())
End Function

Private Sub WaitBeforeRetry(ByVal delayMilliseconds As Long, ByVal Options As HttpExecutionOptions, ByVal totalStarted As Currency)
    Dim waitStarted As Currency
    Dim lastYield As Currency
    Dim elapsed As Double
    Dim remainingWait As Long
    Dim pauseMilliseconds As Long
    Dim beforePause As Currency
    Dim afterPause As Currency

    CheckControls Options, totalStarted
    If delayMilliseconds <= 0 Then Exit Sub
    If Options.TotalDeadlineMilliseconds > 0 And delayMilliseconds >= RemainingDeadline(Options, totalStarted) Then
        Err.Raise HttpErrTimeout, ReliabilitySource, "HTTP total deadline expired before the next retry."
    End If

    waitStarted = mRuntime.CounterValue()
    lastYield = waitStarted
    Do
        CheckControls Options, totalStarted
        elapsed = mRuntime.ElapsedMilliseconds(waitStarted, mRuntime.CounterValue())
        If elapsed >= delayMilliseconds Then Exit Do
        remainingWait = CLng(Fix(CDbl(delayMilliseconds) - elapsed + 0.999999))
        pauseMilliseconds = Options.PollIntervalMilliseconds
        If pauseMilliseconds > remainingWait Then pauseMilliseconds = remainingWait
        If Options.YieldToHost And mRuntime.ElapsedMilliseconds(lastYield, mRuntime.CounterValue()) >= Options.YieldIntervalMilliseconds Then
            mRuntime.YieldHost
            lastYield = mRuntime.CounterValue()
        End If
        beforePause = mRuntime.CounterValue()
        mRuntime.Pause pauseMilliseconds
        afterPause = mRuntime.CounterValue()
        If afterPause <= beforePause Then Err.Raise HttpErrIo, ReliabilitySource, "Reliability runtime clock did not advance during retry wait."
    Loop
End Sub

Private Sub CheckControls(ByVal Options As HttpExecutionOptions, ByVal totalStarted As Currency)
    If Not Options.CancellationToken Is Nothing Then
        If Options.CancellationToken.IsCancellationRequested Then Err.Raise HttpErrCancelled, ReliabilitySource, "HTTP operation was cancelled."
    End If
    If Options.TotalDeadlineMilliseconds > 0 Then
        If mRuntime.ElapsedMilliseconds(totalStarted, mRuntime.CounterValue()) >= Options.TotalDeadlineMilliseconds Then
            Err.Raise HttpErrTimeout, ReliabilitySource, "HTTP total deadline expired."
        End If
    End If
End Sub

Private Function RemainingDeadline(ByVal Options As HttpExecutionOptions, ByVal totalStarted As Currency) As Long
    Dim remaining As Double

    If Options.TotalDeadlineMilliseconds = 0 Then Exit Function
    remaining = CDbl(Options.TotalDeadlineMilliseconds) - mRuntime.ElapsedMilliseconds(totalStarted, mRuntime.CounterValue())
    If remaining <= 0 Then Err.Raise HttpErrTimeout, ReliabilitySource, "HTTP total deadline expired."
    If remaining > 2147483647# Then
        RemainingDeadline = 2147483647
    Else
        RemainingDeadline = CLng(Fix(remaining + 0.999999))
    End If
End Function
