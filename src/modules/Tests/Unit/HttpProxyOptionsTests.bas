Attribute VB_Name = "HttpProxyOptionsTests"
Option Explicit

Public Sub Test_ProxyOptions_DefaultsToSystemProxy()
    Dim options As New HttpProxyOptions

    XlflowAssert.AssertEquals HttpProxyDefault, options.Mode
    XlflowAssert.AssertEquals "", options.ProxyUrl
    XlflowAssert.AssertEquals "", options.BypassList
    XlflowAssert.AssertFalse options.HasOverride
End Sub

Public Sub Test_ProxyOptions_ClonePreservesManualConfiguration()
    Dim options As New HttpProxyOptions
    Dim copy As HttpProxyOptions

    options.Mode = HttpProxyManual
    options.ProxyUrl = "http://127.0.0.1:18080"
    options.BypassList = "localhost;127.0.0.1"
    Set copy = options.Clone()

    XlflowAssert.AssertEquals HttpProxyManual, copy.Mode
    XlflowAssert.AssertEquals "http://127.0.0.1:18080", copy.ProxyUrl
    XlflowAssert.AssertEquals "localhost;127.0.0.1", copy.BypassList
    XlflowAssert.AssertNotSame options, copy
End Sub

Public Sub Test_ProxyOptions_AllowsNoProxyMode()
    Dim options As New HttpProxyOptions

    options.Mode = HttpProxyNoProxy
    options.Validate

    XlflowAssert.AssertTrue options.HasOverride
End Sub

'@ExpectedError(-2147200503, "Manual proxy mode requires a proxy URL.", "HttpProxyOptions")
Public Sub Test_ProxyOptions_RequiresUrlForManualMode()
    Dim options As New HttpProxyOptions

    options.Mode = HttpProxyManual
    options.Validate
End Sub

'@ExpectedError(-2147200503, "No-proxy mode cannot include a manual proxy URL or bypass list.", "HttpProxyOptions")
Public Sub Test_ProxyOptions_RejectsManualFieldsForNoProxy()
    Dim options As New HttpProxyOptions

    options.Mode = HttpProxyNoProxy
    options.ProxyUrl = "http://127.0.0.1:18080"
    options.Validate
End Sub

'@ExpectedError(-2147200503, "Proxy credentials must be configured by the authentication policy, not in the proxy URL.", "HttpProxyOptions.ProxyUrl")
Public Sub Test_ProxyOptions_RejectsCredentialsInUrl()
    Dim options As New HttpProxyOptions

    options.ProxyUrl = "http://" & "proxy-user" & ":" & "proxy-value" & "@127.0.0.1:18080"
End Sub

'@ExpectedError(-2147200503, "Proxy bypass list cannot contain control characters.", "HttpProxyOptions.BypassList")
Public Sub Test_ProxyOptions_RejectsControlCharactersInBypassList()
    Dim options As New HttpProxyOptions

    options.BypassList = "localhost" & vbTab
End Sub
