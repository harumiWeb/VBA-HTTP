Attribute VB_Name = "WinHttpConcurrencyTests"
Option Explicit

Private mScheduledCancellation As HttpCancellationToken

'@Tag("integration")
Public Sub Test_Batch_Concurrency16IsBoundedAndFasterThanSequential()
    Const requestCount As Long = 100
    Dim client As New HttpClient
    Dim Requests As Collection
    Dim Options As New HttpBatchOptions
    Dim sequential As HttpBatchResult
    Dim concurrent As HttpBatchResult
    Dim stats As HttpResponse
    Dim started As Currency
    Dim sequentialMilliseconds As Double
    Dim concurrentMilliseconds As Double

    client.BaseUrl = RequireBaseUrl()
    Options.YieldToHost = False
    Options.PollIntervalMilliseconds = 1
    Options.MaxConcurrency = 1
    Set Requests = DelayRequests(requestCount, 100)
    started = HttpTiming.CounterValue()
    Set sequential = client.ExecuteMany(Requests, Options)
    sequentialMilliseconds = HttpTiming.ElapsedMilliseconds(started, HttpTiming.CounterValue())
    AssertAllSucceeded sequential, requestCount

    Call client.PostResponse("/__admin/reset")
    Options.MaxConcurrency = 16
    Set Requests = DelayRequests(requestCount, 100)
    started = HttpTiming.CounterValue()
    Set concurrent = client.ExecuteMany(Requests, Options)
    concurrentMilliseconds = HttpTiming.ElapsedMilliseconds(started, HttpTiming.CounterValue())
    AssertAllSucceeded concurrent, requestCount

    Set stats = client.GetResponse("/__admin/stats")
    XlflowAssert.AssertTrue concurrentMilliseconds * 6# <= sequentialMilliseconds, "Concurrency 16 must be at least six times faster than sequential execution."
    XlflowAssert.AssertTrue CLng(stats.Headers.GetValue("X-Max-In-Flight")) <= 16, "Server observed more than 16 in-flight requests."
    XlflowAssert.AssertTrue CLng(stats.Headers.GetValue("X-Max-In-Flight")) >= 2, "Batch did not overlap requests."
End Sub

'@Tag("integration")
Public Sub Test_GetMany_PreservesInputOrderAcrossOutOfOrderCompletion()
    Dim client As New HttpClient
    Dim Urls As New Collection
    Dim Options As New HttpBatchOptions
    Dim result As HttpBatchResult

    client.BaseUrl = RequireBaseUrl()
    Urls.Add "/delay/120"
    Urls.Add "/delay/10"
    Urls.Add "/delay/60"
    Options.MaxConcurrency = 3
    Options.PollIntervalMilliseconds = 1
    Options.YieldToHost = False

    Set result = client.GetMany(Urls, Options)

    AssertAllSucceeded result, 3
    XlflowAssert.AssertContains "120", result.ItemAt(1).Response.Text
    XlflowAssert.AssertContains "10", result.ItemAt(2).Response.Text
    XlflowAssert.AssertContains "60", result.ItemAt(3).Response.Text
End Sub

'@Tag("integration")
Public Sub Test_Batch_PreservesPartialFailureAndHttpStatusResponse()
    Dim client As New HttpClient
    Dim Requests As New Collection
    Dim result As HttpBatchResult

    client.BaseUrl = RequireBaseUrl()
    Requests.Add GetRequest("/status/204")
    Requests.Add GetRequest("/disconnect")
    Requests.Add GetRequest("/status/503")

    Set result = client.ExecuteMany(Requests)

    XlflowAssert.AssertEquals 2, result.SuccessCount
    XlflowAssert.AssertEquals 1, result.FailureCount
    XlflowAssert.AssertEquals 0, result.CancelledCount
    XlflowAssert.AssertEquals HttpBatchFailed, result.ItemAt(2).Status
    XlflowAssert.AssertTrue result.ItemAt(2).ErrorCategory <> HttpErrorNone
    XlflowAssert.AssertEquals 503, result.ItemAt(3).Response.StatusCode
End Sub

'@Tag("integration")
Public Sub Test_Batch_RequestDeadlineStopsEveryDelayedOperation()
    Dim client As New HttpClient
    Dim Options As New HttpBatchOptions
    Dim result As HttpBatchResult
    Dim index As Long

    client.BaseUrl = RequireBaseUrl()
    Options.MaxConcurrency = 3
    Options.PollIntervalMilliseconds = 1
    Options.YieldToHost = False
    Options.RequestDeadlineMilliseconds = 25

    Set result = client.ExecuteMany(DelayRequests(3, 250), Options)

    XlflowAssert.AssertEquals 3, result.FailureCount
    For index = 1 To result.Count
        XlflowAssert.AssertEquals HttpErrorTimeout, result.ItemAt(index).ErrorCategory
    Next index
End Sub

'@Tag("integration")
Public Sub Test_Batch_PreCancelledTokenCancelsAllWithoutStartingIo()
    Dim client As New HttpClient
    Dim Options As New HttpBatchOptions
    Dim token As New HttpCancellationToken
    Dim result As HttpBatchResult

    client.BaseUrl = RequireBaseUrl()
    token.Cancel
    Set Options.CancellationToken = token

    Set result = client.ExecuteMany(DelayRequests(4, 100), Options)

    XlflowAssert.AssertEquals 4, result.CancelledCount
    XlflowAssert.AssertEquals 0, result.SuccessCount
    XlflowAssert.AssertEquals 0, result.FailureCount
End Sub

'@Tag("integration")
Public Sub Test_Batch_DoEventsCheckpointObservesRunningCancellation()
    Dim client As New HttpClient
    Dim Options As New HttpBatchOptions
    Dim token As New HttpCancellationToken
    Dim result As HttpBatchResult

    client.BaseUrl = RequireBaseUrl()
    Options.MaxConcurrency = 4
    Options.PollIntervalMilliseconds = 5
    Options.YieldToHost = True
    Options.YieldIntervalMilliseconds = 20
    Set Options.CancellationToken = token
    Set mScheduledCancellation = token
    Application.OnTime Now + TimeSerial(0, 0, 1), "WinHttpConcurrencyTests.CancelScheduledBatch"

    Set result = client.ExecuteMany(DelayRequests(8, 2000), Options)
    Set mScheduledCancellation = Nothing

    XlflowAssert.AssertEquals 8, result.CancelledCount
    XlflowAssert.AssertEquals 0, result.SuccessCount
End Sub

Public Sub CancelScheduledBatch()
    If Not mScheduledCancellation Is Nothing Then mScheduledCancellation.Cancel
End Sub

Private Function DelayRequests(ByVal count As Long, ByVal milliseconds As Long) As Collection
    Dim Requests As New Collection
    Dim index As Long

    For index = 1 To count
        Requests.Add GetRequest("/delay/" & CStr(milliseconds))
    Next index
    Set DelayRequests = Requests
End Function

Private Function GetRequest(ByVal Url As String) As HttpRequest
    Dim Request As New HttpRequest
    Request.Method = "GET"
    Request.Url = Url
    Set GetRequest = Request
End Function

Private Sub AssertAllSucceeded(ByVal result As HttpBatchResult, ByVal expectedCount As Long)
    XlflowAssert.AssertEquals expectedCount, result.Count
    XlflowAssert.AssertEquals expectedCount, result.SuccessCount
    XlflowAssert.AssertEquals 0, result.FailureCount
    XlflowAssert.AssertEquals 0, result.CancelledCount
End Sub

Private Function RequireBaseUrl() As String
    RequireBaseUrl = Trim$(Environ$("VBA_HTTP_TEST_BASE_URL"))
    If Len(RequireBaseUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_BASE_URL is not set; run task test:integration."
        Exit Function
    End If
End Function
