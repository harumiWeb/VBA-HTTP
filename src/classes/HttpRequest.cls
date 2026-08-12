VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMethod As String
Private mUrl As String
Private mHeaders As New HttpHeaders
Private mQuery As New HttpParams
Private mTimeouts As New HttpTimeouts
Private mBody As New HttpBody
Private mFollowRedirects As Boolean
Private mMaxRedirects As Long

Public Property Get Method() As String
    Method = mMethod
End Property

Public Property Let Method(ByVal value As String)
    value = UCase$(Trim$(value))
    Select Case value
    Case "GET", "POST", "PUT", "PATCH", "DELETE", "HEAD", "OPTIONS"
        mMethod = value
    Case Else
        HttpErrors.RaiseValidation "HttpRequest.Method", "Unsupported HTTP method '" & value & "'."
    End Select
End Property

Public Property Get Url() As String
    Url = mUrl
End Property

Public Property Let Url(ByVal value As String)
    If Len(Trim$(value)) = 0 Then
        HttpErrors.RaiseInvalidUrl "HttpRequest.Url", "Request URL cannot be empty."
    End If
    mUrl = value
End Property

Public Property Get Headers() As HttpHeaders
    Set Headers = mHeaders
End Property

Public Property Set Headers(ByVal value As HttpHeaders)
    If value Is Nothing Then
        HttpErrors.RaiseValidation "HttpRequest.Headers", "Headers cannot be Nothing."
    End If
    Set mHeaders = value
End Property

Public Property Get Query() As HttpParams
    Set Query = mQuery
End Property

Public Property Set Query(ByVal value As HttpParams)
    If value Is Nothing Then
        HttpErrors.RaiseValidation "HttpRequest.Query", "Query cannot be Nothing."
    End If
    Set mQuery = value
End Property

Public Property Get Timeouts() As HttpTimeouts
    Set Timeouts = mTimeouts
End Property

Public Property Set Timeouts(ByVal value As HttpTimeouts)
    If value Is Nothing Then
        HttpErrors.RaiseValidation "HttpRequest.Timeouts", "Timeouts cannot be Nothing."
    End If
    Set mTimeouts = value
End Property

Public Property Get Body() As HttpBody
    Set Body = mBody
End Property

Public Property Set Body(ByVal value As HttpBody)
    If value Is Nothing Then
        HttpErrors.RaiseValidation "HttpRequest.Body", "Body cannot be Nothing."
    End If
    Set mBody = value
End Property

Public Property Get FollowRedirects() As Boolean
    FollowRedirects = mFollowRedirects
End Property

Public Property Let FollowRedirects(ByVal value As Boolean)
    mFollowRedirects = value
End Property

Public Property Get MaxRedirects() As Long
    MaxRedirects = mMaxRedirects
End Property

Public Property Let MaxRedirects(ByVal value As Long)
    If value < 1 Or value > 100 Then
        HttpErrors.RaiseValidation "HttpRequest.MaxRedirects", "Maximum redirects must be between 1 and 100."
    End If
    mMaxRedirects = value
End Property

Public Function Clone() As HttpRequest
    Dim copy As New HttpRequest

    copy.Method = mMethod
    If Len(mUrl) > 0 Then copy.Url = mUrl
    Set copy.Headers = mHeaders.Clone()
    Set copy.Query = mQuery.Clone()
    Set copy.Timeouts = mTimeouts.Clone()
    Set copy.Body = mBody.Clone()
    copy.FollowRedirects = mFollowRedirects
    copy.MaxRedirects = mMaxRedirects
    Set Clone = copy
End Function

Private Sub Class_Initialize()
    mMethod = "GET"
    mFollowRedirects = True
    mMaxRedirects = 10
End Sub
