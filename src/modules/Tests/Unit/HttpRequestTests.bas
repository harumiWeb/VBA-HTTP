Attribute VB_Name = "HttpRequestTests"
Option Explicit

Public Sub Test_Request_HasSafeRedirectDefaults()
    Dim Request As New HttpRequest

    XlflowAssert.AssertTrue Request.FollowRedirects
    XlflowAssert.AssertEquals 10, Request.MaxRedirects
End Sub

Public Sub Test_Request_CloneCopiesRedirectPolicy()
    Dim Request As New HttpRequest
    Dim copy As HttpRequest

    Request.FollowRedirects = False
    Request.MaxRedirects = 3
    Set copy = Request.Clone()

    XlflowAssert.AssertFalse copy.FollowRedirects
    XlflowAssert.AssertEquals 3, copy.MaxRedirects
End Sub

'@ExpectedError(-2147200503, "Maximum redirects must be between 1 and 100.", "HttpRequest.MaxRedirects")
Public Sub Test_Request_RejectsZeroRedirectLimit()
    Dim Request As New HttpRequest

    Request.MaxRedirects = 0
End Sub
