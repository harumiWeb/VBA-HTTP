VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpProxyOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMode As HttpProxyMode
Private mProxyUrl As String
Private mBypassList As String

Public Property Get Mode() As HttpProxyMode
    Mode = mMode
End Property

Public Property Let Mode(ByVal value As HttpProxyMode)
    If value <> HttpProxyDefault And value <> HttpProxyNoProxy And value <> HttpProxyManual Then
        HttpErrors.RaiseValidation "HttpProxyOptions.Mode", "Proxy mode must use the default, no-proxy, or manual mode."
    End If
    mMode = value
End Property

Public Property Get ProxyUrl() As String
    ProxyUrl = mProxyUrl
End Property

Public Property Let ProxyUrl(ByVal value As String)
    value = Trim$(value)
    If Len(value) > 0 Then ValidateProxyUrl value
    mProxyUrl = value
End Property

Public Property Get BypassList() As String
    BypassList = mBypassList
End Property

Public Property Let BypassList(ByVal value As String)
    ValidateBypassList value
    mBypassList = Trim$(value)
End Property

Public Property Get HasOverride() As Boolean
    HasOverride = (mMode <> HttpProxyDefault Or Len(mProxyUrl) > 0 Or Len(mBypassList) > 0)
End Property

Public Function Clone() As HttpProxyOptions
    Dim copy As New HttpProxyOptions

    copy.Mode = mMode
    copy.ProxyUrl = mProxyUrl
    copy.BypassList = mBypassList
    copy.Validate
    Set Clone = copy
End Function

Public Sub Validate()
    Select Case mMode
    Case HttpProxyDefault
        If Len(mProxyUrl) > 0 Or Len(mBypassList) > 0 Then
            HttpErrors.RaiseValidation "HttpProxyOptions", "Default proxy mode cannot include a manual proxy URL or bypass list."
        End If
    Case HttpProxyNoProxy
        If Len(mProxyUrl) > 0 Or Len(mBypassList) > 0 Then
            HttpErrors.RaiseValidation "HttpProxyOptions", "No-proxy mode cannot include a manual proxy URL or bypass list."
        End If
    Case HttpProxyManual
        If Len(mProxyUrl) = 0 Then
            HttpErrors.RaiseValidation "HttpProxyOptions", "Manual proxy mode requires a proxy URL."
        End If
        ValidateProxyUrl mProxyUrl
    Case Else
        HttpErrors.RaiseValidation "HttpProxyOptions.Mode", "Proxy mode must use the default, no-proxy, or manual mode."
    End Select
End Sub

Private Sub ValidateProxyUrl(ByVal value As String)
    Dim schemeEnd As Long
    Dim authorityStart As Long
    Dim authorityEnd As Long
    Dim authority As String
    Dim index As Long
    Dim characterCode As Long

    If Len(value) = 0 Then
        HttpErrors.RaiseValidation "HttpProxyOptions.ProxyUrl", "Proxy URL cannot be empty."
    End If
    If LCase$(Left$(value, 7)) <> "http://" And LCase$(Left$(value, 8)) <> "https://" Then
        HttpErrors.RaiseValidation "HttpProxyOptions.ProxyUrl", "Proxy URL must use HTTP or HTTPS."
    End If
    For index = 1 To Len(value)
        characterCode = AscW(Mid$(value, index, 1))
        If characterCode <= 32 Or characterCode = 127 Then
            HttpErrors.RaiseValidation "HttpProxyOptions.ProxyUrl", "Proxy URL cannot contain whitespace or control characters."
        End If
    Next index
    If InStr(1, value, "@", vbBinaryCompare) > 0 Then
        HttpErrors.RaiseValidation "HttpProxyOptions.ProxyUrl", "Proxy credentials must be configured by the authentication policy, not in the proxy URL."
    End If

    schemeEnd = InStr(1, value, "://", vbBinaryCompare)
    authorityStart = schemeEnd + 3
    authorityEnd = Len(value) + 1
    For index = authorityStart To Len(value)
        If InStr(1, "/?#", Mid$(value, index, 1), vbBinaryCompare) > 0 Then
            authorityEnd = index
            Exit For
        End If
    Next index
    authority = Mid$(value, authorityStart, authorityEnd - authorityStart)
    If Len(authority) = 0 Then
        HttpErrors.RaiseValidation "HttpProxyOptions.ProxyUrl", "Proxy URL must contain a host."
    End If
    If authorityEnd <= Len(value) Then
        HttpErrors.RaiseValidation "HttpProxyOptions.ProxyUrl", "Proxy URL cannot contain a path, query, or fragment."
    End If
End Sub

Private Sub ValidateBypassList(ByVal value As String)
    Dim index As Long
    Dim characterCode As Long

    For index = 1 To Len(value)
        characterCode = AscW(Mid$(value, index, 1))
        If characterCode < 32 Or characterCode = 127 Then
            HttpErrors.RaiseValidation "HttpProxyOptions.BypassList", "Proxy bypass list cannot contain control characters."
        End If
    Next index
End Sub

Private Sub Class_Initialize()
    mMode = HttpProxyDefault
    mProxyUrl = ""
    mBypassList = ""
End Sub
