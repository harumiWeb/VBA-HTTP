Attribute VB_Name = "HttpProtocolOptionsTests"
Option Explicit

Public Sub Test_ProtocolOptions_DefaultsToNoOverride()
    Dim options As New HttpProtocolOptions

    XlflowAssert.AssertEquals 0, options.EnabledProtocols
    XlflowAssert.AssertEquals HttpProtocolAllowFallback, options.Mode
    XlflowAssert.AssertFalse options.HasOverride
End Sub

Public Sub Test_ProtocolOptions_ClonePreservesFlagsAndMode()
    Dim options As New HttpProtocolOptions
    Dim copy As HttpProtocolOptions

    options.AllowHttp2 = True
    options.AllowHttp3 = True
    options.Mode = HttpProtocolRequired
    Set copy = options.Clone()

    XlflowAssert.AssertTrue copy.AllowHttp2
    XlflowAssert.AssertTrue copy.AllowHttp3
    XlflowAssert.AssertEquals HttpProtocolRequired, copy.Mode
End Sub

'@ExpectedError(-2147200503, "Required protocol mode needs HTTP/2 or HTTP/3 enabled.", "HttpProtocolOptions")
Public Sub Test_ProtocolOptions_RejectsRequiredModeWithoutProtocols()
    Dim options As New HttpProtocolOptions

    options.Mode = HttpProtocolRequired
    options.Validate
End Sub

'@ExpectedError(-2147200503, "Enabled protocol flags must contain only HTTP/2 and HTTP/3.", "HttpProtocolOptions.EnabledProtocols")
Public Sub Test_ProtocolOptions_RejectsUnknownFlags()
    Dim options As New HttpProtocolOptions

    options.EnabledProtocols = 4
End Sub
