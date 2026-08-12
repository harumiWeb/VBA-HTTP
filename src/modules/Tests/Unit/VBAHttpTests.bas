Attribute VB_Name = "VBAHttpTests"
Option Explicit

Public Sub Test_Factory_ReturnsIndependentDefaultClients()
    Dim first As HttpClient
    Dim second As HttpClient

    Set first = VBAHttp.CreateClient()
    Set second = VBAHttp.CreateClient()
    first.BaseUrl = "http://127.0.0.1:8080"

    XlflowAssert.AssertIsNotNothing first
    XlflowAssert.AssertIsNotNothing second
    XlflowAssert.AssertNotSame first, second
    XlflowAssert.AssertEquals "WinHttpComTransport", TypeName(first.Transport)
    XlflowAssert.AssertEquals "", second.BaseUrl
End Sub
