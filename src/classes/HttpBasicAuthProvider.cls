VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpBasicAuthProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHttpAuthProvider

Private mUsername As String
Private mPassword As String
Private mAllowInsecureHttp As Boolean
Private mInitialized As Boolean

Public Sub Initialize(ByVal Username As String, ByVal Password As String, Optional ByVal AllowInsecureHttp As Boolean = False)
    ValidateUsername Username
    ValidatePassword Password
    mUsername = Username
    mPassword = Password
    mAllowInsecureHttp = AllowInsecureHttp
    mInitialized = True
End Sub

Public Property Get AllowInsecureHttp() As Boolean
    AllowInsecureHttp = mAllowInsecureHttp
End Property

Private Sub IHttpAuthProvider_Apply(ByVal Request As HttpRequest)
    Dim encoded As String

    If Not mInitialized Then
        HttpErrors.RaiseValidation "HttpBasicAuthProvider.Apply", "Basic auth provider is not initialized."
    End If
    ValidateRequest Request
    encoded = HttpEncoding.EncodeBase64(HttpEncoding.EncodeUtf8(mUsername & ":" & mPassword))
    Request.Headers.SetValue "Authorization", "Basic " & encoded
End Sub

Private Function IHttpAuthProvider_Clone() As IHttpAuthProvider
    Dim copy As New HttpBasicAuthProvider

    If mInitialized Then copy.Initialize mUsername, mPassword, mAllowInsecureHttp
    Set IHttpAuthProvider_Clone = copy
End Function

Private Sub ValidateRequest(ByVal Request As HttpRequest)
    Dim scheme As String

    If Request Is Nothing Then
        HttpErrors.RaiseValidation "HttpBasicAuthProvider.Apply", "Request cannot be Nothing."
    End If
    If Request.Headers.Contains("Authorization") Then
        HttpErrors.RaiseValidation "HttpBasicAuthProvider.Apply", "Authorization header conflicts with the configured auth provider."
    End If
    scheme = LCase$(Left$(Request.Url, 8))
    If scheme = "https://" Then Exit Sub
    If Left$(scheme, 7) = "http://" Then
        If Not mAllowInsecureHttp Then HttpErrors.RaiseValidation "HttpBasicAuthProvider.Apply", "Basic authentication requires an HTTPS request URL."
        Exit Sub
    End If
    If scheme <> "https://" Then
        HttpErrors.RaiseValidation "HttpBasicAuthProvider.Apply", "Basic authentication requires an HTTP or HTTPS request URL."
    End If
End Sub

Private Sub ValidateUsername(ByVal value As String)
    If Len(value) = 0 Then
        HttpErrors.RaiseValidation "HttpBasicAuthProvider.Initialize", "Basic username cannot be empty."
    End If
    If InStr(1, value, ":", vbBinaryCompare) > 0 Then
        HttpErrors.RaiseValidation "HttpBasicAuthProvider.Initialize", "Basic username cannot contain a colon."
    End If
    ValidateCredentialText value, "username"
End Sub

Private Sub ValidatePassword(ByVal value As String)
    ValidateCredentialText value, "password"
End Sub

Private Sub ValidateCredentialText(ByVal value As String, ByVal label As String)
    Dim index As Long
    Dim characterCode As Long

    For index = 1 To Len(value)
        characterCode = AscW(Mid$(value, index, 1))
        If characterCode < 32 Or characterCode = 127 Then
            HttpErrors.RaiseValidation "HttpBasicAuthProvider.Initialize", "Basic " & label & " cannot contain control characters."
        End If
    Next index
End Sub
