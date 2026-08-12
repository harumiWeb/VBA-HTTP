Attribute VB_Name = "WinHttpUploadStressTests"
Option Explicit

'@Tag("stress")
Public Sub Test_NativeTransport_UploadsOneGiBFile()
    Dim client As New HttpClient
    Dim options As New HttpExecutionOptions
    Dim result As HttpUploadResult
    Dim sourcePath As String
    Dim expectedHash As String
    Dim expectedBytes As Currency

    sourcePath = Trim$(Environ$("VBA_HTTP_UPLOAD_STRESS_PATH"))
    If Len(sourcePath) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_UPLOAD_STRESS_PATH is not set; run task test:upload-stress."
        Exit Sub
    End If
    If Len(Trim$(Environ$("VBA_HTTP_TEST_BASE_URL"))) = 0 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_TEST_BASE_URL is not set; run task test:upload-stress."
        Exit Sub
    End If
    expectedHash = LCase$(Trim$(Environ$("VBA_HTTP_UPLOAD_STRESS_EXPECTED_HASH")))
    If Len(expectedHash) <> 64 Then
        XlflowAssert.AssertInconclusive "VBA_HTTP_UPLOAD_STRESS_EXPECTED_HASH is not set; run task test:upload-stress."
        Exit Sub
    End If

    expectedBytes = CCur(1073741824#)
    WaitForUploadStressGate
    client.BaseUrl = Trim$(Environ$("VBA_HTTP_TEST_BASE_URL"))
    Set client.Transport = New WinHttpNativeTransport
    options.TotalDeadlineMilliseconds = 1800000
    options.YieldToHost = False
    Set result = client.UploadFile("/upload/hash", sourcePath, "application/octet-stream", options)

    XlflowAssert.AssertEquals 200, result.StatusCode
    XlflowAssert.AssertEquals expectedBytes, result.BytesWritten
    XlflowAssert.AssertEquals expectedBytes, result.ContentLength
    XlflowAssert.AssertEquals expectedHash, LCase$(result.Headers.GetValue("X-Upload-Digest"))
    XlflowAssert.AssertEquals "1073741824", result.Headers.GetValue("X-Upload-Bytes")
End Sub

Private Sub WaitForUploadStressGate()
    Dim fileSystem As Object
    Dim goPath As String
    Dim deadline As Date

    goPath = Trim$(Environ$("VBA_HTTP_UPLOAD_STRESS_GO_PATH"))
    If Len(goPath) = 0 Then
        XlflowAssert.AssertInconclusive "Stress gate paths are not configured; run task test:upload-stress."
        Exit Sub
    End If
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    deadline = DateAdd("s", 30, Now)
    Do While Not fileSystem.FileExists(goPath)
        If Now >= deadline Then Err.Raise HttpErrTimeout, "WinHttpUploadStressTests", "Stress memory gate was not released."
        HttpTiming.Pause 50
    Loop
End Sub
