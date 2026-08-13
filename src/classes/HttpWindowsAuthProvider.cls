VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpWindowsAuthProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHttpAuthProvider

Private Const ProviderSource As String = "HttpWindowsAuthProvider"
Private Const DefaultChallengeLimit As Long = 3

Private mUsername As String
Private mPassword As String
Private mScheme As HttpAuthChallengeScheme
Private mTarget As HttpAuthChallengeTarget
Private mMaxChallenges As Long
Private mAllowInsecureHttp As Boolean
Private mInitialized As Boolean

Public Sub Initialize(ByVal Username As String, ByVal Password As String, Optional ByVal Scheme As HttpAuthChallengeScheme = HttpAuthSchemeAuto, Optional ByVal Target As HttpAuthChallengeTarget = HttpAuthTargetServer, Optional ByVal AllowInsecureHttp As Boolean = False, Optional ByVal MaxChallenges As Long = DefaultChallengeLimit)
    ValidateCredentials Username, Password
    ValidateScheme Scheme
    ValidateTarget Target
    If MaxChallenges < 1 Or MaxChallenges > 3 Then
        HttpErrors.RaiseValidation ProviderSource & ".Initialize", "Maximum authentication challenges must be between 1 and 3."
    End If
    mUsername = Username
    mPassword = Password
    mScheme = Scheme
    mTarget = Target
    mMaxChallenges = MaxChallenges
    mAllowInsecureHttp = AllowInsecureHttp
    mInitialized = True
End Sub

Public Property Get Scheme() As HttpAuthChallengeScheme
    Scheme = mScheme
End Property

Public Property Get Target() As HttpAuthChallengeTarget
    Target = mTarget
End Property

Public Property Get MaxChallenges() As Long
    MaxChallenges = mMaxChallenges
End Property

Public Property Get AllowInsecureHttp() As Boolean
    AllowInsecureHttp = mAllowInsecureHttp
End Property

Private Sub IHttpAuthProvider_Apply(ByVal Request As HttpRequest)
    If Not mInitialized Then HttpErrors.RaiseValidation ProviderSource & ".Apply", "Windows authentication provider is not initialized."
    ValidateRequest Request
    Request.FollowRedirects = False
End Sub

Private Function IHttpAuthProvider_Clone() As IHttpAuthProvider
    Dim copy As New HttpWindowsAuthProvider

    If mInitialized Then copy.Initialize mUsername, mPassword, mScheme, mTarget, mAllowInsecureHttp, mMaxChallenges
    Set IHttpAuthProvider_Clone = copy
End Function

Friend Sub ConfigureCom(ByVal Backend As Object, ByVal Source As String)
    Dim targetBackend As Object

    If Not mInitialized Then HttpErrors.RaiseValidation Source, "Windows authentication provider is not initialized."
    ValidateComCompatibility Source
    Set targetBackend = Backend
    If targetBackend Is Nothing Then
        HttpErrors.RaiseValidation Source, "WinHTTP backend is unavailable."
    Else
        On Error GoTo ConfigureFailed
        targetBackend.SetCredentials mUsername, mPassword, CLng(mTarget)
    End If
    Exit Sub

ConfigureFailed:
    Err.Raise HttpErrIo, Source, "WinHTTP rejected the configured authentication credentials."
End Sub

Friend Sub ValidateComCompatibility(ByVal Source As String)
    If Not mInitialized Then HttpErrors.RaiseValidation Source, "Windows authentication provider is not initialized."
    If mScheme <> HttpAuthSchemeAuto Then
        HttpErrors.RaiseValidation Source, "Explicit challenge scheme selection requires the native WinHTTP transport; use HttpAuthSchemeAuto for COM."
    End If
End Sub

#If VBA7 Then
Friend Function ApplyNativeCredentials(ByVal HandleValue As LongPtr, ByVal ChallengeTarget As Long, ByVal ChallengeScheme As Long, ByVal Source As String) As Boolean
    If Not mInitialized Then HttpErrors.RaiseValidation Source, "Windows authentication provider is not initialized."
    If ChallengeTarget <> CLng(mTarget) Then Exit Function
    If Not IsSupportedScheme(ChallengeScheme) Then Exit Function
    ApplyNativeCredentials = WinHttpNativeApi.SetCredentials(HandleValue, ChallengeTarget, ChallengeScheme, mUsername, mPassword)
End Function
#Else
Friend Function ApplyNativeCredentials(ByVal HandleValue As Long, ByVal ChallengeTarget As Long, ByVal ChallengeScheme As Long, ByVal Source As String) As Boolean
    If Not mInitialized Then HttpErrors.RaiseValidation Source, "Windows authentication provider is not initialized."
    If ChallengeTarget <> CLng(mTarget) Then Exit Function
    If Not IsSupportedScheme(ChallengeScheme) Then Exit Function
    ApplyNativeCredentials = WinHttpNativeApi.SetCredentials(HandleValue, ChallengeTarget, ChallengeScheme, mUsername, mPassword)
End Function
#End If

Friend Function SelectScheme(ByVal SupportedSchemes As Long, ByVal FirstScheme As Long, ByVal ChallengeTarget As Long) As Long
    If Not mInitialized Then Exit Function
    If ChallengeTarget <> CLng(mTarget) Then Exit Function
    If mScheme <> HttpAuthSchemeAuto Then
        If (SupportedSchemes And CLng(mScheme)) <> 0 Then SelectScheme = CLng(mScheme)
        Exit Function
    End If
    If (SupportedSchemes And CLng(HttpAuthSchemeNegotiate)) <> 0 Then
        SelectScheme = CLng(HttpAuthSchemeNegotiate)
    ElseIf (SupportedSchemes And CLng(HttpAuthSchemeNtlm)) <> 0 Then
        SelectScheme = CLng(HttpAuthSchemeNtlm)
    ElseIf (SupportedSchemes And CLng(HttpAuthSchemeDigest)) <> 0 Then
        SelectScheme = CLng(HttpAuthSchemeDigest)
    ElseIf (SupportedSchemes And CLng(HttpAuthSchemeBasic)) <> 0 Then
        SelectScheme = CLng(HttpAuthSchemeBasic)
    ElseIf IsSupportedScheme(FirstScheme) Then
        SelectScheme = FirstScheme
    End If
End Function

Private Function IsSupportedScheme(ByVal SchemeValue As Long) As Boolean
    IsSupportedScheme = (SchemeValue = HttpAuthSchemeBasic Or SchemeValue = HttpAuthSchemeNtlm Or SchemeValue = HttpAuthSchemeDigest Or SchemeValue = HttpAuthSchemeNegotiate)
End Function

Private Sub ValidateRequest(ByVal Request As HttpRequest)
    Dim scheme As String
    Dim credentialHeader As String

    If Request Is Nothing Then HttpErrors.RaiseValidation ProviderSource & ".Apply", "Request cannot be Nothing."
    credentialHeader = HttpSecurity.CredentialHeaderName(mTarget)
    If Request.Headers.Contains(credentialHeader) Then
        HttpErrors.RaiseValidation ProviderSource & ".Apply", credentialHeader & " header conflicts with the configured challenge provider."
    End If
    scheme = LCase$(Left$(Request.Url, 8))
    If scheme = "https://" Then Exit Sub
    If Left$(scheme, 7) = "http://" Then
        If Not mAllowInsecureHttp Then HttpErrors.RaiseValidation ProviderSource & ".Apply", "Challenge authentication requires an HTTPS request URL."
        Exit Sub
    End If
    HttpErrors.RaiseValidation ProviderSource & ".Apply", "Challenge authentication requires an HTTP or HTTPS request URL."
End Sub

Private Sub ValidateCredentials(ByVal Username As String, ByVal Password As String)
    If Len(Username) = 0 Then HttpErrors.RaiseValidation ProviderSource & ".Initialize", "Authentication username cannot be empty."
    ValidateCredentialText Username, "username"
    ValidateCredentialText Password, "password"
End Sub

Private Sub ValidateCredentialText(ByVal Value As String, ByVal Label As String)
    Dim index As Long
    Dim characterCode As Long

    For index = 1 To Len(Value)
        characterCode = AscW(Mid$(Value, index, 1))
        If characterCode < 32 Or characterCode = 127 Then
            HttpErrors.RaiseValidation ProviderSource & ".Initialize", "Authentication " & Label & " cannot contain control characters."
        End If
    Next index
End Sub

Private Sub ValidateScheme(ByVal SchemeValue As HttpAuthChallengeScheme)
    If SchemeValue <> HttpAuthSchemeAuto And SchemeValue <> HttpAuthSchemeBasic And SchemeValue <> HttpAuthSchemeNtlm And SchemeValue <> HttpAuthSchemeDigest And SchemeValue <> HttpAuthSchemeNegotiate Then
        HttpErrors.RaiseValidation ProviderSource & ".Initialize", "Unsupported authentication scheme."
    End If
End Sub

Private Sub ValidateTarget(ByVal TargetValue As HttpAuthChallengeTarget)
    If TargetValue <> HttpAuthTargetServer And TargetValue <> HttpAuthTargetProxy Then
        HttpErrors.RaiseValidation ProviderSource & ".Initialize", "Unsupported authentication target."
    End If
End Sub
