Attribute VB_Name = "HttpBatchTests"
Option Explicit

Public Sub Test_BatchOptions_HasBoundedDefaultsAndIndependentClone()
    Dim Options As New HttpBatchOptions
    Dim copy As HttpBatchOptions
    Dim token As New HttpCancellationToken

    Set Options.CancellationToken = token
    Set copy = Options.Clone()
    copy.MaxConcurrency = 4

    XlflowAssert.AssertEquals 8, Options.MaxConcurrency
    XlflowAssert.AssertEquals 5, Options.PollIntervalMilliseconds
    XlflowAssert.AssertTrue Options.YieldToHost
    XlflowAssert.AssertEquals 20, Options.YieldIntervalMilliseconds
    XlflowAssert.AssertEquals 0, Options.RequestDeadlineMilliseconds
    XlflowAssert.AssertEquals 4, copy.MaxConcurrency
    XlflowAssert.AssertSame token, copy.CancellationToken
End Sub

'@ExpectedError(-2147200503, "MaxConcurrency must be between 1 and 64.", "HttpBatchOptions.MaxConcurrency")
Public Sub Test_BatchOptions_RejectsZeroConcurrency()
    Dim Options As New HttpBatchOptions
    Options.MaxConcurrency = 0
End Sub

Public Sub Test_BatchResult_PreservesOrderAndCountsTerminalStates()
    Dim result As New HttpBatchResult
    Dim succeeded As New HttpBatchItem
    Dim failed As New HttpBatchItem
    Dim cancelled As New HttpBatchItem
    Dim response As New HttpResponse

    response.Initialize 204
    succeeded.InitializeSuccess 1, response
    failed.InitializeFailure 2, HttpErrTimeout, "test", "timed out"
    cancelled.InitializeCancelled 3
    result.Add succeeded
    result.Add failed
    result.Add cancelled

    XlflowAssert.AssertEquals 3, result.Count
    XlflowAssert.AssertEquals 1, result.SuccessCount
    XlflowAssert.AssertEquals 1, result.FailureCount
    XlflowAssert.AssertEquals 1, result.CancelledCount
    XlflowAssert.AssertEquals 2, result.ItemAt(2).Index
    XlflowAssert.AssertEquals HttpErrorTimeout, result.ItemAt(2).ErrorCategory
End Sub

Public Sub Test_CancellationToken_IsMonotonic()
    Dim token As New HttpCancellationToken

    XlflowAssert.AssertFalse token.IsCancellationRequested
    token.Cancel
    token.Cancel
    XlflowAssert.AssertTrue token.IsCancellationRequested
End Sub

Public Sub Test_ExecuteMany_EmptyInputReturnsEmptyResult()
    Dim client As New HttpClient
    Dim Requests As New Collection
    Dim result As HttpBatchResult

    Set result = client.ExecuteMany(Requests)

    XlflowAssert.AssertEquals 0, result.Count
End Sub

'@ExpectedError(-2147200503, "Configured transport does not support batch execution.", "HttpClient.ExecuteMany")
Public Sub Test_ExecuteMany_RejectsTransportWithoutBatchCapability()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim Requests As New Collection

    Set client.Transport = transport
    Call client.ExecuteMany(Requests)
End Sub

'@ExpectedError(-2147200503, "The same HttpClient cannot execute reentrantly.", "HttpClient.Execute")
Public Sub Test_ExecuteMany_RejectsSameClientReentrancy()
    Dim client As New HttpClient
    Dim transport As New ReentrantBatchTransport
    Dim Requests As New Collection
    Dim Request As New HttpRequest

    Request.Url = "http://127.0.0.1:1/status/204"
    Requests.Add Request
    Set transport.Client = client
    Set client.Transport = transport
    Call client.ExecuteMany(Requests)
End Sub

Public Sub Test_BatchWorkItem_TracksRetryAndTerminalState()
    Dim work As New HttpBatchWorkItem
    Dim Request As New HttpRequest
    Dim response As New HttpResponse

    Request.Url = "http://127.0.0.1/status/204"
    work.Initialize 2, Request
    XlflowAssert.AssertTrue work.IsEligible(0)
    work.BeginAttempt
    XlflowAssert.AssertEquals 1, work.Attempt
    XlflowAssert.AssertTrue work.IsRunning
    work.ScheduleRetry 10, 50
    XlflowAssert.AssertFalse work.IsEligible(59)
    XlflowAssert.AssertTrue work.IsEligible(60)
    work.BeginAttempt
    response.Initialize 204
    work.CompleteSuccess response
    XlflowAssert.AssertTrue work.IsComplete
    XlflowAssert.AssertEquals 2, work.Attempt
    XlflowAssert.AssertEquals 2, work.TerminalItem.Index
End Sub
