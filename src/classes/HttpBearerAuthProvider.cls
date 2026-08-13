VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpBearerAuthProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHttpAuthProvider

Private mToken As String
Private mAllowInsecureHttp As Boolean
Private mInitialized As Boolean

Public Sub Initialize(ByVal Token As String, Optional ByVal AllowInsecureHttp As Boolean = False)
    ValidateToken Token
    mToken = Token
    mAllowInsecureHttp = AllowInsecureHttp
    mInitialized = True
End Sub

Public Property Get AllowInsecureHttp() As Boolean
    AllowInsecureHttp = mAllowInsecureHttp
End Property

Private Sub IHttpAuthProvider_Apply(ByVal Request As HttpRequest)
    If Not mInitialized Then
        HttpErrors.RaiseValidation "HttpBearerAuthProvider.Apply", "Bearer auth provider is not initialized."
    End If
    ValidateRequest Request
    Request.Headers.SetValue "Authorization", "Bearer " & mToken
End Sub

Private Function IHttpAuthProvider_Clone() As IHttpAuthProvider
    Dim copy As New HttpBearerAuthProvider

    If mInitialized Then copy.Initialize mToken, mAllowInsecureHttp
    Set IHttpAuthProvider_Clone = copy
End Function

Private Sub ValidateRequest(ByVal Request As HttpRequest)
    Dim scheme As String

    If Request Is Nothing Then
        HttpErrors.RaiseValidation "HttpBearerAuthProvider.Apply", "Request cannot be Nothing."
    End If
    If Request.Headers.Contains("Authorization") Then
        HttpErrors.RaiseValidation "HttpBearerAuthProvider.Apply", "Authorization header conflicts with the configured auth provider."
    End If
    scheme = LCase$(Left$(Request.Url, 8))
    If scheme = "https://" Then Exit Sub
    If Left$(scheme, 7) = "http://" Then
        If Not mAllowInsecureHttp Then HttpErrors.RaiseValidation "HttpBearerAuthProvider.Apply", "Bearer authentication requires an HTTPS request URL."
        Exit Sub
    End If
    If scheme <> "https://" Then
        HttpErrors.RaiseValidation "HttpBearerAuthProvider.Apply", "Bearer authentication requires an HTTP or HTTPS request URL."
    End If
End Sub

Private Sub ValidateToken(ByVal value As String)
    Dim index As Long
    Dim characterCode As Long

    If Len(value) = 0 Then
        HttpErrors.RaiseValidation "HttpBearerAuthProvider.Initialize", "Bearer token cannot be empty."
    End If
    For index = 1 To Len(value)
        characterCode = AscW(Mid$(value, index, 1))
        If characterCode < 0 Then characterCode = characterCode + 65536
        If characterCode <= 32 Or characterCode = 127 Then
            HttpErrors.RaiseValidation "HttpBearerAuthProvider.Initialize", "Bearer token cannot contain whitespace or control characters."
        End If
        If characterCode > 127 Then
            HttpErrors.RaiseValidation "HttpBearerAuthProvider.Initialize", "Bearer token must contain ASCII characters only."
        End If
    Next index
End Sub
