VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpCookieRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mName As String
Private mValue As String
Private mDomain As String
Private mPath As String
Private mHostOnly As Boolean
Private mSecure As Boolean
Private mHasExpires As Boolean
Private mExpires As Date

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property

Public Property Get Value() As String
    Value = mValue
End Property

Public Property Let Value(ByVal value As String)
    mValue = value
End Property

Public Property Get Domain() As String
    Domain = mDomain
End Property

Public Property Let Domain(ByVal value As String)
    mDomain = LCase$(value)
End Property

Public Property Get Path() As String
    Path = mPath
End Property

Public Property Let Path(ByVal value As String)
    mPath = value
End Property

Public Property Get HostOnly() As Boolean
    HostOnly = mHostOnly
End Property

Public Property Let HostOnly(ByVal value As Boolean)
    mHostOnly = value
End Property

Public Property Get Secure() As Boolean
    Secure = mSecure
End Property

Public Property Let Secure(ByVal value As Boolean)
    mSecure = value
End Property

Public Property Get HasExpires() As Boolean
    HasExpires = mHasExpires
End Property

Public Property Let HasExpires(ByVal value As Boolean)
    mHasExpires = value
End Property

Public Property Get ExpiresAt() As Date
    ExpiresAt = mExpires
End Property

Public Property Let ExpiresAt(ByVal value As Date)
    mExpires = value
End Property

Public Property Get SerializedLength() As Long
    SerializedLength = LenB(StrConv(mName & "=" & mValue & "; Domain=" & mDomain & "; Path=" & mPath, vbFromUnicode))
End Property

Public Function IsExpired(ByVal nowUtc As Date) As Boolean
    IsExpired = mHasExpires And mExpires <= nowUtc
End Function

Public Function Matches(ByVal target As HttpCookieUrl, ByVal nowUtc As Date) As Boolean
    If IsExpired(nowUtc) Then Exit Function
    If mSecure And Not target.IsHttps Then Exit Function
    If mHostOnly Then
        If StrComp(target.Host, mDomain, vbTextCompare) <> 0 Then Exit Function
    ElseIf Not HttpCookieParsing.DomainMatches(target.Host, mDomain) Then
        Exit Function
    End If
    Matches = HttpCookieParsing.PathMatches(target.Path, mPath)
End Function
