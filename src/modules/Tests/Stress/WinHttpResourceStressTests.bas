Attribute VB_Name = "WinHttpResourceStressTests"
Option Explicit

Private Const SequentialHandleDeltaLimit As Long = 8
Private Const ScheduledHandleDeltaLimit As Long = 32
Private Const WarmupRequests As Long = 10000

'@Tag("stress")
Public Sub Test_ResourceStress_SequentialNative()
    Dim client As New HttpClient
    Dim response As HttpResponse
    Dim baseUrl As String
    Dim startMarker As String
    Dim doneMarker As String
    Dim releaseMarker As String
    Dim iterations As Long
    Dim beforeHandles As Long
    Dim afterHandles As Long
    Dim index As Long

    baseUrl = RequireEnvironment("VBA_HTTP_TEST_BASE_URL")
    startMarker = RequireEnvironment("VBA_HTTP_RESOURCE_START_PATH")
    doneMarker = RequireEnvironment("VBA_HTTP_RESOURCE_DONE_PATH")
    releaseMarker = RequireEnvironment("VBA_HTTP_RESOURCE_RELEASE_PATH")
    iterations = ReadBoundedLong("VBA_HTTP_RESOURCE_ITERATIONS", 10000, 1, 100000)

    client.BaseUrl = baseUrl
    Set client.Transport = New WinHttpNativeTransport
    For index = 1 To WarmupRequests
        Set response = client.GetResponse("/status/204")
        XlflowAssert.AssertEquals 204, response.StatusCode
    Next index
    Set response = Nothing

    If Not WinHttpNativeApi.ProcessHandleCount(beforeHandles) Then
        XlflowAssert.AssertInconclusive "GetProcessHandleCount is unavailable on this host."
        Exit Sub
    End If
    WriteMarker startMarker

    For index = 1 To iterations
        Set response = client.GetResponse("/status/204")
        XlflowAssert.AssertEquals 204, response.StatusCode
    Next index
    Set response = Nothing

    If Not WinHttpNativeApi.ProcessHandleCount(afterHandles) Then
        XlflowAssert.AssertInconclusive "GetProcessHandleCount is unavailable after sequential stress."
        Exit Sub
    End If
    XlflowAssert.AssertTrue afterHandles - beforeHandles <= SequentialHandleDeltaLimit, "Sequential native request handles grew persistently."
    WriteMarker doneMarker
    WaitForRelease releaseMarker, "WinHttpResourceStressTests.SequentialNative"
End Sub

'@Tag("stress")
Public Sub Test_ResourceStress_ScheduledCom()
    Dim client As New HttpClient
    Dim Requests As New Collection
    Dim Options As New HttpBatchOptions
    Dim result As HttpBatchResult
    Dim Request As HttpRequest
    Dim baseUrl As String
    Dim startMarker As String
    Dim doneMarker As String
    Dim releaseMarker As String
    Dim iterations As Long
    Dim concurrency As Long
    Dim beforeHandles As Long
    Dim afterHandles As Long
    Dim index As Long

    baseUrl = RequireEnvironment("VBA_HTTP_TEST_BASE_URL")
    startMarker = RequireEnvironment("VBA_HTTP_RESOURCE_START_PATH")
    doneMarker = RequireEnvironment("VBA_HTTP_RESOURCE_DONE_PATH")
    releaseMarker = RequireEnvironment("VBA_HTTP_RESOURCE_RELEASE_PATH")
    iterations = ReadBoundedLong("VBA_HTTP_RESOURCE_ITERATIONS", 10000, 1, 100000)
    concurrency = ReadBoundedLong("VBA_HTTP_RESOURCE_CONCURRENCY", 16, 1, 64)

    client.BaseUrl = baseUrl
    For index = 1 To WarmupRequests
        Set Request = New HttpRequest
        Request.Method = "GET"
        Request.Url = "/status/204"
        Requests.Add Request
    Next index
    Options.MaxConcurrency = concurrency
    Options.PollIntervalMilliseconds = 1
    Options.YieldToHost = False
    Set result = client.ExecuteMany(Requests, Options)
    XlflowAssert.AssertEquals WarmupRequests, result.Count
    XlflowAssert.AssertEquals WarmupRequests, result.SuccessCount
    XlflowAssert.AssertEquals 0, result.FailureCount
    Set result = Nothing
    Set Requests = New Collection
    For index = 1 To iterations
        Set Request = New HttpRequest
        Request.Method = "GET"
        Request.Url = "/status/204"
        Requests.Add Request
    Next index
    If Not WinHttpNativeApi.ProcessHandleCount(beforeHandles) Then
        XlflowAssert.AssertInconclusive "GetProcessHandleCount is unavailable on this host."
        Exit Sub
    End If
    WriteMarker startMarker

    Set result = client.ExecuteMany(Requests, Options)
    XlflowAssert.AssertEquals iterations, result.Count
    XlflowAssert.AssertEquals iterations, result.SuccessCount
    XlflowAssert.AssertEquals 0, result.FailureCount
    XlflowAssert.AssertEquals 0, result.CancelledCount

    If Not WinHttpNativeApi.ProcessHandleCount(afterHandles) Then
        XlflowAssert.AssertInconclusive "GetProcessHandleCount is unavailable after scheduled stress."
        Exit Sub
    End If
    XlflowAssert.AssertTrue afterHandles - beforeHandles <= ScheduledHandleDeltaLimit, "Scheduled COM request handles grew persistently."
    WriteMarker doneMarker
    WaitForRelease releaseMarker, "WinHttpResourceStressTests.ScheduledCom"
End Sub

Private Function RequireEnvironment(ByVal Name As String) As String
    RequireEnvironment = Trim$(Environ$(Name))
    If Len(RequireEnvironment) = 0 Then
        XlflowAssert.AssertInconclusive Name & " is not set; run task test:resource-stress."
    End If
End Function

Private Function ReadBoundedLong(ByVal Name As String, ByVal defaultValue As Long, ByVal minimum As Long, ByVal maximum As Long) As Long
    Dim raw As String

    raw = Trim$(Environ$(Name))
    If Len(raw) = 0 Then
        ReadBoundedLong = defaultValue
        Exit Function
    End If
    If Not IsNumeric(raw) Then HttpErrors.RaiseValidation "WinHttpResourceStressTests", Name & " must be numeric."
    On Error GoTo InvalidValue
    ReadBoundedLong = CLng(raw)
    On Error GoTo 0
    If ReadBoundedLong < minimum Or ReadBoundedLong > maximum Then HttpErrors.RaiseValidation "WinHttpResourceStressTests", Name & " is outside the supported range."
    Exit Function

InvalidValue:
    On Error GoTo 0
    HttpErrors.RaiseValidation "WinHttpResourceStressTests", Name & " is not a valid 32-bit integer."
End Function

Private Sub WriteMarker(ByVal path As String)
    Dim fileSystem As Object
    Dim stream As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem Is Nothing Then Err.Raise HttpErrIo, "WinHttpResourceStressTests.WriteMarker", "File-system marker support is unavailable."
    Set stream = CallByName(fileSystem, "CreateTextFile", VbMethod, path, True, False)
    If stream Is Nothing Then Err.Raise HttpErrIo, "WinHttpResourceStressTests.WriteMarker", "Could not create the resource stress marker."
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
        If Now >= deadline Then Err.Raise HttpErrTimeout, Source, "Resource stress release gate was not opened."
        HttpTiming.Pause 50
    Loop
    Set fileSystem = Nothing
End Sub
