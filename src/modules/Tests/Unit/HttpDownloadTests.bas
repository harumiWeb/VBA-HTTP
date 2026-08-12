Attribute VB_Name = "HttpDownloadTests"
Option Explicit

Public Sub Test_DownloadResult_PreservesLargeByteCounts()
    Dim Headers As New HttpHeaders
    Dim result As New HttpDownloadResult
    Dim largeCount As Currency

    largeCount = CCur(2147483648#)
    Headers.Add "Content-Length", "2147483648"
    result.Initialize 200, "OK", Headers, "HTTP/1.1", "C:\Temp\large.bin", largeCount, True, 12.5
    result.MarkPublished largeCount

    XlflowAssert.AssertTrue result.IsSuccess
    XlflowAssert.AssertTrue result.Published
    XlflowAssert.AssertEquals largeCount, result.BytesWritten
    XlflowAssert.AssertEquals largeCount, result.ContentLength
    XlflowAssert.AssertTrue result.ContentLengthKnown
End Sub

Public Sub Test_DownloadResult_UnknownLengthUsesNegativeOne()
    Dim Headers As New HttpHeaders
    Dim result As New HttpDownloadResult

    result.Initialize 200, "OK", Headers, "HTTP/1.1", "C:\Temp\unknown.bin", 0, False, 1

    XlflowAssert.AssertFalse result.Published
    XlflowAssert.AssertFalse result.ContentLengthKnown
    XlflowAssert.AssertEquals - 1, result.ContentLength
End Sub

'@ExpectedError(-2147200494, "HTTP request returned status 404.", "HttpDownloadResult.RaiseForStatus")
Public Sub Test_DownloadResult_NonSuccessCanBeRaisedExplicitly()
    Dim Headers As New HttpHeaders
    Dim result As New HttpDownloadResult

    result.Initialize 404, "Not Found", Headers, "HTTP/1.1", "C:\Temp\missing.bin", 0, False, 1
    result.RaiseForStatus
End Sub
