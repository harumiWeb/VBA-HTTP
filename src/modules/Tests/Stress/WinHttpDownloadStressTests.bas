Attribute VB_Name = "WinHttpDownloadStressTests"
Option Explicit

'@Tag("stress")
Public Sub Test_NativeTransport_DownloadsOneGiBStream()
    Dim client As New HttpClient
    Dim options As New HttpExecutionOptions
    Dim result As HttpDownloadResult
    Dim destination As String
    Dim warmupDestination As String
    Dim expectedBytes As Currency

    destination = Trim$(Environ$("VBA_HTTP_DOWNLOAD_STRESS_PATH"))
    If Len(destination) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_DOWNLOAD_STRESS_PATH is not set; run task test:download-stress."
        Exit Sub
    End If
    If Len(Trim$(Environ$("VBA_HTTP_TEST_BASE_URL"))) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_BASE_URL is not set; run task test:download-stress."
        Exit Sub
    End If

    expectedBytes = CCur(1073741824#)
    client.BaseUrl = Trim$(Environ$("VBA_HTTP_TEST_BASE_URL"))
    Set client.Transport = New WinHttpNativeTransport
    warmupDestination = destination & ".warmup"
    Set result = client.DownloadFile("/bytes/65536", warmupDestination, options)
    XlflowAssert.AssertTrue result.Published
    DeleteDownloadFile warmupDestination
    WaitForDownloadStressGate
    options.TotalDeadlineMilliseconds = 900000
    options.YieldToHost = False
    Set result = client.DownloadFile("/stream/1073741824", destination, options)

    XlflowAssert.AssertEquals 200, result.StatusCode
    XlflowAssert.AssertTrue result.Published
    XlflowAssert.AssertEquals expectedBytes, result.BytesWritten
    XlflowAssert.AssertEquals expectedBytes, FileLen(destination)
End Sub

Private Sub WaitForDownloadStressGate()
    Dim fileSystem As Object
    Dim goPath As String
    Dim deadline As Date

    goPath = Trim$(Environ$("VBA_HTTP_DOWNLOAD_STRESS_GO_PATH"))
    If Len(goPath) = 0 Then
        XlflowAssert.AssertInconclusive "Stress gate paths are not configured; run task test:download-stress."
        Exit Sub
    End If
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    deadline = DateAdd("s", 30, Now)
    Do While Not fileSystem.FileExists(goPath)
        If Now >= deadline Then Err.Raise HttpErrTimeout, "WinHttpDownloadStressTests", "Stress memory gate was not released."
        HttpTiming.Pause 50
    Loop
End Sub
