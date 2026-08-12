Attribute VB_Name = "WinHttpReliabilityTests"
Option Explicit

Private mScheduledCancellation As HttpCancellationToken

'@Tag("integration")
Public Sub Test_Retry_FlakyGetSucceedsOnThirdAttempt()
    Dim client As New HttpClient
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set response = client.GetResponse("/flaky/2?id=reliability-flaky")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertContains """attempt"":3", response.Text
End Sub

'@Tag("integration")
Public Sub Test_Retry_RateLimitHonorsZeroRetryAfter()
    Dim client As New HttpClient
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set response = client.GetResponse("/rate-limit/1?id=reliability-rate&retry_after=0")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertContains """attempt"":2", response.Text
End Sub

'@Tag("integration")
'@TestCase("408"; 408)
'@TestCase("429"; 429)
'@TestCase("500"; 500)
'@TestCase("502"; 502)
'@TestCase("503"; 503)
'@TestCase("504"; 504)
Public Sub Test_Retry_DefaultStatusMatrix(ByVal statusCode As Long)
    Dim client As New HttpClient
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set response = client.GetResponse("/retry-status/" & CStr(statusCode) & "/1?id=status-" & CStr(statusCode))

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertContains """attempt"":2", response.Text
End Sub

'@Tag("integration")
Public Sub Test_Retry_PostDoesNotRetryWithoutOptIn()
    Dim client As New HttpClient
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    Set response = client.PostResponse("/retry-status/503/1?id=post-default")

    XlflowAssert.AssertEquals 503, response.StatusCode
    XlflowAssert.AssertContains """attempt"":1", response.Text
End Sub

'@Tag("integration")
Public Sub Test_Retry_PostRetriesWithExplicitOptIn()
    Dim client As New HttpClient
    Dim policy As New HttpRetryPolicy
    Dim response As HttpResponse

    client.BaseUrl = RequireBaseUrl()
    policy.RetryNonIdempotentMethods = True
    policy.BaseDelayMilliseconds = 0
    policy.MaxDelayMilliseconds = 0
    Set client.RetryPolicy = policy
    Set response = client.PostResponse("/retry-status/503/1?id=post-opt-in")

    XlflowAssert.AssertEquals 200, response.StatusCode
    XlflowAssert.AssertContains """attempt"":2", response.Text
End Sub

'@Tag("integration")
'@ExpectedError(-2147200498)
Public Sub Test_Reliability_ActiveAttemptStopsAtTotalDeadline()
    Dim client As New HttpClient
    Dim Options As New HttpExecutionOptions
    Dim policy As New HttpRetryPolicy

    client.BaseUrl = RequireBaseUrl()
    Options.TotalDeadlineMilliseconds = 100
    Options.YieldToHost = False
    policy.MaxAttempts = 1
    Set Options.RetryPolicy = policy
    Call client.GetResponse("/delay/2000", Nothing, Options)
End Sub

'@Tag("integration")
'@ExpectedError(-2147200497)
Public Sub Test_Reliability_ActiveAttemptObservesCancellation()
    Dim client As New HttpClient
    Dim Options As New HttpExecutionOptions
    Dim policy As New HttpRetryPolicy
    Dim token As New HttpCancellationToken

    client.BaseUrl = RequireBaseUrl()
    Options.YieldToHost = True
    Options.YieldIntervalMilliseconds = 20
    Set Options.CancellationToken = token
    policy.MaxAttempts = 1
    Set Options.RetryPolicy = policy
    Set mScheduledCancellation = token
    Application.OnTime Now + TimeSerial(0, 0, 1), "WinHttpReliabilityTests.CancelScheduledRequest"

    On Error GoTo ExpectedFailure
    Call client.GetResponse("/delay/2000", Nothing, Options)
    Set mScheduledCancellation = Nothing
    Exit Sub

ExpectedFailure:
    Set mScheduledCancellation = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub CancelScheduledRequest()
    If Not mScheduledCancellation Is Nothing Then mScheduledCancellation.Cancel
End Sub

Private Function RequireBaseUrl() As String
    RequireBaseUrl = Trim$(Environ$("VBA_HTTP_TEST_BASE_URL"))
    If Len(RequireBaseUrl) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_BASE_URL is not set; run task test:integration."
        Exit Function
    End If
End Function
