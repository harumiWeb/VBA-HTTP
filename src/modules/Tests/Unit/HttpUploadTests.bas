Attribute VB_Name = "HttpUploadTests"
Option Explicit

Public Sub Test_UploadResult_ReportsChallengeAndLargeCounts()
    Dim Headers As New HttpHeaders
    Dim result As New HttpUploadResult
    Dim largeCount As Currency

    largeCount = CCur(2147483648#)
    Headers.Add "WWW-Authenticate", "Basic realm=""test"""
    result.Initialize 401, "Unauthorized", Headers, "HTTP/1.1", largeCount, largeCount, 4.5

    XlflowAssert.AssertFalse result.IsSuccess
    XlflowAssert.AssertTrue result.AuthenticationChallenged
    XlflowAssert.AssertEquals largeCount, result.BytesWritten
    XlflowAssert.AssertEquals largeCount, result.ContentLength
End Sub

Public Sub Test_MultipartForm_PreservesOrderAndUtf8Length()
    Dim form As New HttpMultipartForm

    form.Boundary = "----test-boundary"
    form.AddField "title", "日本語"
    form.AddField "empty", ""

    XlflowAssert.AssertEquals "----test-boundary", form.Boundary
    XlflowAssert.AssertEquals 2, form.Count
    XlflowAssert.AssertEquals "title", form.PartAt(1).Name
    XlflowAssert.AssertEquals "日本語", form.PartAt(1).Value
    XlflowAssert.AssertEquals 9, WinHttpUploadEncoding.Utf8ByteCount("日本語")
End Sub

'@ExpectedError(-2147200503, "Boundary contains an unsafe character.", "HttpMultipartForm.Boundary")
Public Sub Test_MultipartForm_RejectsHeaderInjectionBoundary()
    Dim form As New HttpMultipartForm

    form.Boundary = "bad" & vbCrLf & "boundary"
End Sub

'@ExpectedError(-2147200503, "Multipart header token contains an unsafe character.", "HttpMultipartForm.AddField")
Public Sub Test_MultipartForm_RejectsHeaderInjectionName()
    Dim form As New HttpMultipartForm

    form.AddField "field" & vbCrLf & "X-Evil: yes", "value"
End Sub
