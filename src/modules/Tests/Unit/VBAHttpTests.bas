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

Public Sub Test_Factory_CreatesReliabilityConfigurationObjects()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateRetryPolicy()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateExecutionOptions()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateBatchOptions()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateCancellationToken()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateProtocolOptions()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateDecompressionOptions()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateProxyOptions()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateCookieJar()
End Sub

Public Sub Test_Factory_CreatesAuthenticationProviders()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateBasicAuthProvider("user", "pass")
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateBearerAuthProvider("token", True)
End Sub
