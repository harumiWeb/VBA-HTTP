VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpProtocolOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEnabledProtocols As Long
Private mMode As HttpProtocolMode

Public Property Get EnabledProtocols() As Long
    EnabledProtocols = mEnabledProtocols
End Property

Public Property Let EnabledProtocols(ByVal value As Long)
    If value < 0 Or value > (HttpProtocolHttp2 Or HttpProtocolHttp3) Then
        HttpErrors.RaiseValidation "HttpProtocolOptions.EnabledProtocols", "Enabled protocol flags must contain only HTTP/2 and HTTP/3."
    End If
    mEnabledProtocols = value
End Property

Public Property Get Mode() As HttpProtocolMode
    Mode = mMode
End Property

Public Property Let Mode(ByVal value As HttpProtocolMode)
    If value <> HttpProtocolAllowFallback And value <> HttpProtocolRequired Then
        HttpErrors.RaiseValidation "HttpProtocolOptions.Mode", "Protocol mode must allow fallback or require the selected protocols."
    End If
    mMode = value
End Property

Public Property Get HasOverride() As Boolean
    HasOverride = (mEnabledProtocols <> 0)
End Property

Public Property Get AllowHttp2() As Boolean
    AllowHttp2 = ((mEnabledProtocols And HttpProtocolHttp2) <> 0)
End Property

Public Property Let AllowHttp2(ByVal value As Boolean)
    If value Then
        mEnabledProtocols = mEnabledProtocols Or HttpProtocolHttp2
    Else
        mEnabledProtocols = mEnabledProtocols And Not HttpProtocolHttp2
    End If
End Property

Public Property Get AllowHttp3() As Boolean
    AllowHttp3 = ((mEnabledProtocols And HttpProtocolHttp3) <> 0)
End Property

Public Property Let AllowHttp3(ByVal value As Boolean)
    If value Then
        mEnabledProtocols = mEnabledProtocols Or HttpProtocolHttp3
    Else
        mEnabledProtocols = mEnabledProtocols And Not HttpProtocolHttp3
    End If
End Property

Public Function Clone() As HttpProtocolOptions
    Dim copy As New HttpProtocolOptions

    copy.EnabledProtocols = mEnabledProtocols
    copy.Mode = mMode
    copy.Validate
    Set Clone = copy
End Function

Public Sub Validate()
    If mMode = HttpProtocolRequired And mEnabledProtocols = 0 Then
        HttpErrors.RaiseValidation "HttpProtocolOptions", "Required protocol mode needs HTTP/2 or HTTP/3 enabled."
    End If
End Sub

Private Sub Class_Initialize()
    mEnabledProtocols = 0
    mMode = HttpProtocolAllowFallback
End Sub
