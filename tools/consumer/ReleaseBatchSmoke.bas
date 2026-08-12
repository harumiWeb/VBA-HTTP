Attribute VB_Name = "ReleaseBatchSmoke"
Option Explicit

Public Sub RunBatchSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim Urls As New Collection
    Dim result As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    client.BaseUrl = baseUrl
    Urls.Add "/delay/100"
    Urls.Add "/delay/100"
    Urls.Add "/delay/100"
    Urls.Add "/delay/100"
    Set result = client.GetMany(Urls)

    If result.Count <> 4 Or result.SuccessCount <> 4 Or result.FailureCount <> 0 Or result.CancelledCount <> 0 Then
        Err.Raise vbObjectError + 740, "ReleaseBatchSmoke.RunBatchSmoke", "Release batch smoke returned an unexpected result."
    End If
End Sub

Public Sub RunReliabilitySmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim policy As Object
    Dim Options As Object
    Dim response As Object
    Dim errorNumber As Long

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    Set policy = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateRetryPolicy")
    policy.BaseDelayMilliseconds = 0
    policy.MaxDelayMilliseconds = 0
    Set client.RetryPolicy = policy
    client.BaseUrl = baseUrl

    Set response = client.GetResponse("/flaky/2?id=release-reliability")
    If response.StatusCode <> 200 Or InStr(1, response.Text, """attempt"":3", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 741, "ReleaseBatchSmoke.RunReliabilitySmoke", "Release retry smoke returned an unexpected result."
    End If

    Set Options = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateExecutionOptions")
    Options.TotalDeadlineMilliseconds = 100
    Options.YieldToHost = False
    On Error Resume Next
    Set response = client.GetResponse("/delay/2000", Nothing, Options)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    If errorNumber <> -2147200498 Then
        Err.Raise vbObjectError + 742, "ReleaseBatchSmoke.RunReliabilitySmoke", "Release total-deadline smoke did not return HttpErrTimeout."
    End If
End Sub
