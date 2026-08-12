VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpNativeUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mScheme As String
Private mHost As String
Private mPath As String
Private mPort As Long
Private mSecure As Boolean

Friend Property Get Scheme() As String
    Scheme = mScheme
End Property

Friend Property Get Host() As String
    Host = mHost
End Property

Friend Property Get PathAndQuery() As String
    PathAndQuery = mPath
End Property

Friend Property Get Port() As Long
    Port = mPort
End Property

Friend Property Get IsSecure() As Boolean
    IsSecure = mSecure
End Property

Friend Sub Initialize(ByVal Url As String)
    Dim schemeSeparator As Long
    Dim authorityStart As Long
    Dim authorityEnd As Long
    Dim authority As String
    Dim rest As String
    Dim colonIndex As Long

    If Len(Url) = 0 Or Trim$(Url) <> Url Then
        HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "URL must be an absolute HTTP or HTTPS URL."
    End If
    schemeSeparator = InStr(1, Url, "://", vbBinaryCompare)
    If schemeSeparator <= 1 Then
        HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "URL must be an absolute HTTP or HTTPS URL."
    End If
    mScheme = LCase$(Left$(Url, schemeSeparator - 1))
    If mScheme <> "http" And mScheme <> "https" Then
        HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "URL must use HTTP or HTTPS."
    End If
    mSecure = (mScheme = "https")
    authorityStart = schemeSeparator + 3
    authorityEnd = FindAuthorityEnd(Url, authorityStart)
    authority = Mid$(Url, authorityStart, authorityEnd - authorityStart)
    If Len(authority) = 0 Or InStr(1, authority, "@", vbBinaryCompare) > 0 Then
        HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "URL authority is invalid."
    End If

    ParseAuthority authority, mHost, mPort
    ValidateHost mHost
    If mPort = 0 Then
        If mSecure Then
            mPort = 443
        Else
            mPort = 80
        End If
    End If

    If authorityEnd > Len(Url) Then
        mPath = "/"
    Else
        rest = Mid$(Url, authorityEnd)
        colonIndex = InStr(1, rest, "#", vbBinaryCompare)
        If colonIndex > 0 Then rest = Left$(rest, colonIndex - 1)
        If Len(rest) = 0 Then
            mPath = "/"
        ElseIf Left$(rest, 1) = "?" Then
            mPath = "/" & rest
        ElseIf Left$(rest, 1) = "/" Then
            mPath = rest
        Else
            HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "URL path is invalid."
        End If
    End If
End Sub

Private Function FindAuthorityEnd(ByVal Url As String, ByVal StartIndex As Long) As Long
    Dim index As Long

    For index = StartIndex To Len(Url)
        If InStr(1, "/?#", Mid$(Url, index, 1), vbBinaryCompare) > 0 Then
            FindAuthorityEnd = index
            Exit Function
        End If
    Next index
    FindAuthorityEnd = Len(Url) + 1
End Function

Private Sub ParseAuthority(ByVal authority As String, ByRef host As String, ByRef port As Long)
    Dim closingBracket As Long
    Dim colonIndex As Long
    Dim portText As String

    host = authority
    port = 0
    If Left$(authority, 1) = "[" Then
        closingBracket = InStr(2, authority, "]", vbBinaryCompare)
        If closingBracket <= 2 Then HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "IPv6 authority is invalid."
        host = Mid$(authority, 2, closingBracket - 2)
        If closingBracket < Len(authority) Then
            If Mid$(authority, closingBracket + 1, 1) <> ":" Then HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "IPv6 port separator is invalid."
            portText = Mid$(authority, closingBracket + 2)
            port = ParsePort(portText)
        End If
        Exit Sub
    End If

    colonIndex = InStr(1, authority, ":", vbBinaryCompare)
    If colonIndex > 0 And colonIndex = InStrRev(authority, ":") Then
        host = Left$(authority, colonIndex - 1)
        portText = Mid$(authority, colonIndex + 1)
        port = ParsePort(portText)
    End If
End Sub

Private Function ParsePort(ByVal value As String) As Long
    Dim parsed As Long

    If Len(value) = 0 Or Not IsDigits(value) Then HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "URL port is invalid."
    On Error GoTo InvalidPort
    parsed = CLng(value)
    On Error GoTo 0
    If parsed < 1 Or parsed > 65535 Then GoTo InvalidPort
    ParsePort = parsed
    Exit Function
InvalidPort:
    On Error GoTo 0
    HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "URL port is outside the valid range."
End Function

Private Function IsDigits(ByVal value As String) As Boolean
    Dim index As Long
    Dim code As Long

    For index = 1 To Len(value)
        code = AscW(Mid$(value, index, 1))
        If code < 48 Or code > 57 Then Exit Function
    Next index
    IsDigits = (Len(value) > 0)
End Function

Private Sub ValidateHost(ByVal host As String)
    Dim index As Long
    Dim code As Long

    If Len(host) = 0 Then HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "URL must contain a host."
    For index = 1 To Len(host)
        code = AscW(Mid$(host, index, 1))
        If code <= 32 Or code = 127 Then HttpErrors.RaiseInvalidUrl "WinHttpNativeUrl.Initialize", "URL host contains whitespace or a control character."
    Next index
End Sub
