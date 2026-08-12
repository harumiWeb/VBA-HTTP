VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpDecompressionOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEnabledEncodings As Long
Private mMode As HttpDecompressionMode

Public Property Get EnabledEncodings() As Long
    EnabledEncodings = mEnabledEncodings
End Property

Public Property Let EnabledEncodings(ByVal value As Long)
    If value < 0 Or value > (HttpDecompressionGzip Or HttpDecompressionDeflate) Then
        HttpErrors.RaiseValidation "HttpDecompressionOptions.EnabledEncodings", "Enabled decompression flags must contain only gzip and deflate."
    End If
    mEnabledEncodings = value
End Property

Public Property Get Mode() As HttpDecompressionMode
    Mode = mMode
End Property

Public Property Let Mode(ByVal value As HttpDecompressionMode)
    If value <> HttpDecompressionAllowFallback And value <> HttpDecompressionRequired Then
        HttpErrors.RaiseValidation "HttpDecompressionOptions.Mode", "Decompression mode must allow fallback or require the selected encodings."
    End If
    mMode = value
End Property

Public Property Get HasOverride() As Boolean
    HasOverride = (mEnabledEncodings <> 0)
End Property

Public Property Get AllowGzip() As Boolean
    AllowGzip = ((mEnabledEncodings And HttpDecompressionGzip) <> 0)
End Property

Public Property Let AllowGzip(ByVal value As Boolean)
    If value Then
        mEnabledEncodings = mEnabledEncodings Or HttpDecompressionGzip
    Else
        mEnabledEncodings = mEnabledEncodings And Not HttpDecompressionGzip
    End If
End Property

Public Property Get AllowDeflate() As Boolean
    AllowDeflate = ((mEnabledEncodings And HttpDecompressionDeflate) <> 0)
End Property

Public Property Let AllowDeflate(ByVal value As Boolean)
    If value Then
        mEnabledEncodings = mEnabledEncodings Or HttpDecompressionDeflate
    Else
        mEnabledEncodings = mEnabledEncodings And Not HttpDecompressionDeflate
    End If
End Property

Public Function Clone() As HttpDecompressionOptions
    Dim copy As New HttpDecompressionOptions

    copy.EnabledEncodings = mEnabledEncodings
    copy.Mode = mMode
    copy.Validate
    Set Clone = copy
End Function

Public Sub Validate()
    If mMode = HttpDecompressionRequired And mEnabledEncodings = 0 Then
        HttpErrors.RaiseValidation "HttpDecompressionOptions", "Required decompression mode needs gzip or deflate enabled."
    End If
End Sub

Private Sub Class_Initialize()
    mEnabledEncodings = 0
    mMode = HttpDecompressionAllowFallback
End Sub
