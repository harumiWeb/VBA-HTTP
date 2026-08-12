Attribute VB_Name = "HttpRetryPolicyTests"
Option Explicit

Public Sub Test_RetryPolicy_HasSafeDefaults()
    Dim policy As New HttpRetryPolicy

    XlflowAssert.AssertEquals 3, policy.MaxAttempts
    XlflowAssert.AssertEquals 200, policy.BaseDelayMilliseconds
    XlflowAssert.AssertEquals 30000, policy.MaxDelayMilliseconds
    XlflowAssert.AssertNear 0.2, policy.JitterRatio, 0.000001
    XlflowAssert.AssertTrue policy.RespectRetryAfter
    XlflowAssert.AssertFalse policy.RetryNonIdempotentMethods
End Sub

Public Sub Test_RetryPolicy_AllowsOnlyDefaultIdempotentMethods()
    Dim policy As New HttpRetryPolicy

    XlflowAssert.AssertTrue policy.AllowsMethod("GET")
    XlflowAssert.AssertTrue policy.AllowsMethod("put")
    XlflowAssert.AssertTrue policy.AllowsMethod("DELETE")
    XlflowAssert.AssertFalse policy.AllowsMethod("POST")
    XlflowAssert.AssertFalse policy.AllowsMethod("PATCH")
    XlflowAssert.AssertFalse policy.AllowsMethod("CUSTOM")

    policy.RetryNonIdempotentMethods = True
    XlflowAssert.AssertTrue policy.AllowsMethod("POST")
    XlflowAssert.AssertTrue policy.AllowsMethod("CUSTOM")
End Sub

Public Sub Test_RetryPolicy_ClassifiesDefaultStatusesAndErrors()
    Dim policy As New HttpRetryPolicy

    XlflowAssert.AssertTrue policy.AllowsStatus(408)
    XlflowAssert.AssertTrue policy.AllowsStatus(429)
    XlflowAssert.AssertTrue policy.AllowsStatus(500)
    XlflowAssert.AssertTrue policy.AllowsStatus(502)
    XlflowAssert.AssertTrue policy.AllowsStatus(503)
    XlflowAssert.AssertTrue policy.AllowsStatus(504)
    XlflowAssert.AssertFalse policy.AllowsStatus(400)
    XlflowAssert.AssertFalse policy.AllowsStatus(501)

    XlflowAssert.AssertTrue policy.AllowsError(HttpErrorDns)
    XlflowAssert.AssertTrue policy.AllowsError(HttpErrorConnection)
    XlflowAssert.AssertTrue policy.AllowsError(HttpErrorTimeout)
    XlflowAssert.AssertTrue policy.AllowsError(HttpErrorIo)
    XlflowAssert.AssertFalse policy.AllowsError(HttpErrorTls)
    XlflowAssert.AssertFalse policy.AllowsError(HttpErrorCancelled)
    XlflowAssert.AssertFalse policy.AllowsError(HttpErrorProtocol)
End Sub

Public Sub Test_RetryPolicy_ComputesCappedSymmetricJitter()
    Dim policy As New HttpRetryPolicy

    policy.JitterRatio = 0.5
    policy.MaxDelayMilliseconds = 500

    XlflowAssert.AssertEquals 100, policy.BackoffMilliseconds(1, 0)
    XlflowAssert.AssertEquals 200, policy.BackoffMilliseconds(1, 0.5)
    XlflowAssert.AssertEquals 300, policy.BackoffMilliseconds(1, 1)
    XlflowAssert.AssertEquals 500, policy.BackoffMilliseconds(3, 1)
    XlflowAssert.AssertEquals 500, policy.BackoffMilliseconds(100, 1)
End Sub

'@ExpectedError(-2147200503, "Maximum delay cannot be less than base delay.", "HttpRetryPolicy")
Public Sub Test_RetryPolicy_RejectsInvalidDelayRelationship()
    Dim policy As New HttpRetryPolicy

    policy.MaxDelayMilliseconds = 100
    Call policy.Clone()
End Sub

Public Sub Test_ExecutionOptions_CloneIsolatesPolicyAndSharesCancellation()
    Dim Options As New HttpExecutionOptions
    Dim policy As New HttpRetryPolicy
    Dim token As New HttpCancellationToken
    Dim copy As HttpExecutionOptions

    policy.MaxAttempts = 5
    Set Options.RetryPolicy = policy
    Set Options.CancellationToken = token
    Options.TotalDeadlineMilliseconds = 1000
    Set copy = Options.Clone()

    policy.MaxAttempts = 9
    copy.RetryPolicy.MaxAttempts = 2
    XlflowAssert.AssertEquals 5, Options.RetryPolicy.MaxAttempts
    XlflowAssert.AssertEquals 2, copy.RetryPolicy.MaxAttempts
    XlflowAssert.AssertSame token, copy.CancellationToken
    XlflowAssert.AssertEquals 1000, copy.TotalDeadlineMilliseconds
End Sub

Public Sub Test_BatchOptions_CloneIncludesReliabilityControls()
    Dim Options As New HttpBatchOptions
    Dim policy As New HttpRetryPolicy
    Dim copy As HttpBatchOptions

    policy.MaxAttempts = 4
    Set Options.RetryPolicy = policy
    Options.TotalDeadlineMilliseconds = 2500
    Set copy = Options.Clone()

    XlflowAssert.AssertEquals 4, copy.RetryPolicy.MaxAttempts
    XlflowAssert.AssertEquals 2500, copy.TotalDeadlineMilliseconds
End Sub
