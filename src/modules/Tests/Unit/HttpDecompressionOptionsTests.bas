Attribute VB_Name = "HttpDecompressionOptionsTests"
Option Explicit

Public Sub Test_DecompressionOptions_DefaultsToNoOverride()
    Dim options As New HttpDecompressionOptions

    XlflowAssert.AssertEquals 0, options.EnabledEncodings
    XlflowAssert.AssertEquals HttpDecompressionAllowFallback, options.Mode
    XlflowAssert.AssertFalse options.HasOverride
End Sub

Public Sub Test_DecompressionOptions_ClonePreservesFlagsAndMode()
    Dim options As New HttpDecompressionOptions
    Dim copy As HttpDecompressionOptions

    options.AllowGzip = True
    options.AllowDeflate = True
    options.Mode = HttpDecompressionRequired
    Set copy = options.Clone()

    XlflowAssert.AssertTrue copy.AllowGzip
    XlflowAssert.AssertTrue copy.AllowDeflate
    XlflowAssert.AssertEquals HttpDecompressionRequired, copy.Mode
End Sub

'@ExpectedError(-2147200503, "Required decompression mode needs gzip or deflate enabled.", "HttpDecompressionOptions")
Public Sub Test_DecompressionOptions_RejectsRequiredModeWithoutEncodings()
    Dim options As New HttpDecompressionOptions

    options.Mode = HttpDecompressionRequired
    options.Validate
End Sub

'@ExpectedError(-2147200503, "Enabled decompression flags must contain only gzip and deflate.", "HttpDecompressionOptions.EnabledEncodings")
Public Sub Test_DecompressionOptions_RejectsUnknownFlags()
    Dim options As New HttpDecompressionOptions

    options.EnabledEncodings = 4
End Sub
