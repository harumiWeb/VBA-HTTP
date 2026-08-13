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
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateDiagnostics()
End Sub

Public Sub Test_Factory_CreatesIndependentRequestsWithTimeouts()
    Dim first As HttpRequest
    Dim second As HttpRequest

    Set first = VBAHttp.CreateRequest()
    Set second = VBAHttp.CreateRequest()
    first.Timeouts.ReceiveMilliseconds = 15000

    XlflowAssert.AssertNotSame first, second
    XlflowAssert.AssertEquals 15000, first.Timeouts.ReceiveMilliseconds
    XlflowAssert.AssertEquals 300000, second.Timeouts.ReceiveMilliseconds
End Sub

Public Sub Test_Factory_CreatesAuthenticationProviders()
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateBasicAuthProvider("user", "pass")
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateBearerAuthProvider("token", True)
    XlflowAssert.AssertIsNotNothing VBAHttp.CreateWindowsAuthProvider("user", "pass", HttpAuthSchemeBasic, HttpAuthTargetServer, True)
End Sub
