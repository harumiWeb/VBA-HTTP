Attribute VB_Name = "HttpReliabilityEngineTests"
Option Explicit

Public Sub Test_SequencedTransport_ReturnsConfiguredResponse()
    Dim transportObject As New SequencedHttpTransport
    Dim transport As IHttpTransport
    Dim Request As New HttpRequest
    Dim response As HttpResponse

    transportObject.EnqueueResponse 503
    Request.Url = "http://127.0.0.1/status"
    Set transport = transportObject
    Set response = transport.Execute(Request)

    XlflowAssert.AssertEquals 503, response.StatusCode
End Sub

Public Sub Test_Reliability_RetriesIdempotentStatusWithBackoff()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim response As HttpResponse

    transport.EnqueueResponse 503
    transport.EnqueueResponse 200
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime

    Set response = client.GetResponse("http://127.0.0.1/status")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals 2, transport.ExecuteCount
    XlflowAssert.AssertEquals 200, runtime.PausedMilliseconds
End Sub

Public Sub Test_Reliability_DoesNotRetryPostWithoutOptIn()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim response As HttpResponse

    transport.EnqueueResponse 503
    transport.EnqueueResponse 200
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime

    Set response = client.PostResponse("http://127.0.0.1/status")

    XlflowAssert.AssertEquals 503, response.StatusCode
    XlflowAssert.AssertEquals 1, transport.ExecuteCount
    XlflowAssert.AssertEquals 0, runtime.PausedMilliseconds
End Sub

Public Sub Test_Reliability_RetriesPostWithExplicitOptIn()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim policy As New HttpRetryPolicy
    Dim response As HttpResponse

    transport.EnqueueResponse 503
    transport.EnqueueResponse 200
    policy.RetryNonIdempotentMethods = True
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime
    Set client.RetryPolicy = policy

    Set response = client.PostResponse("http://127.0.0.1/status")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals 2, transport.ExecuteCount
End Sub

Public Sub Test_Reliability_RetriesConnectionFailureAndRaisesFinalError()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim errorNumber As Long

    transport.EnqueueFailure HttpErrConnection, "connection one"
    transport.EnqueueFailure HttpErrConnection, "connection two"
    transport.EnqueueFailure HttpErrConnection, "connection three"
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime

    On Error Resume Next
    Call client.GetResponse("http://127.0.0.1/status")
    errorNumber = Err.Number
    On Error GoTo 0

    XlflowAssert.AssertEquals HttpErrConnection, errorNumber
    XlflowAssert.AssertEquals 3, transport.ExecuteCount
    XlflowAssert.AssertEquals 600, runtime.PausedMilliseconds
End Sub

Public Sub Test_Reliability_DoesNotRetryTlsFailure()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim errorNumber As Long

    transport.EnqueueFailure HttpErrTls, "certificate rejected"
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime

    On Error Resume Next
    Call client.GetResponse("https://127.0.0.1/status")
    errorNumber = Err.Number
    On Error GoTo 0

    XlflowAssert.AssertEquals HttpErrTls, errorNumber
    XlflowAssert.AssertEquals 1, transport.ExecuteCount
End Sub

Public Sub Test_Reliability_UsesRetryAfterDeltaWithoutJitter()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim response As HttpResponse

    transport.EnqueueResponse 429, "2"
    transport.EnqueueResponse 200
    runtime.JitterUnit = 0
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime

    Set response = client.GetResponse("http://127.0.0.1/status")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals 2000, runtime.PausedMilliseconds
End Sub

Public Sub Test_Reliability_UsesRetryAfterHttpDate()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim response As HttpResponse

    transport.EnqueueResponse 503, "Thu, 01 Jan 2026 00:00:03 GMT"
    transport.EnqueueResponse 200
    runtime.UtcNow = DateSerial(2026, 1, 1)
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime

    Set response = client.GetResponse("http://127.0.0.1/status")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertEquals 3000, runtime.PausedMilliseconds
End Sub

Public Sub Test_RetryAfter_ParsesAllHttpDateFormsAndSaturates()
    Dim delay As Long
    Dim nowUtc As Date

    nowUtc = DateSerial(1994, 11, 6) + TimeSerial(8, 49, 35)
    XlflowAssert.AssertTrue HttpRetryAfter.TryParse("Sun, 06 Nov 1994 08:49:37 GMT", nowUtc, delay), "IMF-fixdate"
    XlflowAssert.AssertEquals 2000, delay
    XlflowAssert.AssertTrue HttpRetryAfter.TryParse("Sunday, 06-Nov-94 08:49:37 GMT", nowUtc, delay), "RFC850"
    XlflowAssert.AssertEquals 2000, delay
    XlflowAssert.AssertTrue HttpRetryAfter.TryParse("Sun Nov  6 08:49:37 1994", nowUtc, delay), "asctime"
    XlflowAssert.AssertEquals 2000, delay
    XlflowAssert.AssertTrue HttpRetryAfter.TryParse("999999999999", nowUtc, delay), "delta saturation"
    XlflowAssert.AssertEquals 2147483647, delay
    XlflowAssert.AssertFalse HttpRetryAfter.TryParse("not-a-date", nowUtc, delay)
End Sub

'@ExpectedError(-2147200498, "HTTP total deadline expired before the next retry.", "HttpClient.Execute")
Public Sub Test_Reliability_TotalDeadlineStopsRetryWait()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim Options As New HttpExecutionOptions

    transport.EnqueueResponse 503
    Options.TotalDeadlineMilliseconds = 100
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime
    Call client.GetResponse("http://127.0.0.1/status", Nothing, Options)
End Sub

'@ExpectedError(-2147200497, "HTTP operation was cancelled.", "HttpClient.Execute")
Public Sub Test_Reliability_CancellationStopsRetryWait()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim Options As New HttpExecutionOptions
    Dim token As New HttpCancellationToken

    transport.EnqueueResponse 503
    Set Options.CancellationToken = token
    runtime.CancelAfter 50, token
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime
    Call client.GetResponse("http://127.0.0.1/status", Nothing, Options)
End Sub

'@ExpectedError(-2147200503, "Configured transport cannot enforce cancellation or total deadline during an active attempt.", "HttpClient.Execute")
Public Sub Test_Reliability_RejectsControlForBlockingCustomTransport()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim Options As New HttpExecutionOptions

    Options.TotalDeadlineMilliseconds = 100
    Set client.Transport = transport
    Call client.GetResponse("http://127.0.0.1/status", Nothing, Options)
End Sub

'@ExpectedError(-2147200497, "HTTP operation was cancelled.", "HttpClient.Execute")
Public Sub Test_Reliability_CancellationTakesPriorityOverExpiredDeadline()
    Dim client As New HttpClient
    Dim transport As New SequencedHttpTransport
    Dim runtime As New FakeHttpReliabilityRuntime
    Dim Options As New HttpExecutionOptions
    Dim token As New HttpCancellationToken

    transport.EnqueueResponse 200
    token.Cancel
    Set Options.CancellationToken = token
    Options.TotalDeadlineMilliseconds = 1
    Set client.Transport = transport
    Set client.ReliabilityRuntime = runtime
    Call client.GetResponse("http://127.0.0.1/status", Nothing, Options)
End Sub
