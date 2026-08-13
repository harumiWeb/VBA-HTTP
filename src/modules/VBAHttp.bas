Attribute VB_Name = "VBAHttp"
Option Explicit

''' Creates a client with the default buffered WinHTTP transport.
'''
''' Returns:
'''     A new independently configured HttpClient instance.
Public Function CreateClient() As HttpClient
    Set CreateClient = New HttpClient
End Function

''' Creates a client using the synchronous native WinHTTP transport.
Public Function CreateNativeClient() As HttpClient
    Dim client As New HttpClient

    Set client.Transport = New WinHttpNativeTransport
    Set CreateNativeClient = client
End Function

''' Creates a caller-owned request snapshot for referenced-workbook consumers.
'''
''' PublicNotCreatable domain classes cannot be constructed directly across a
''' workbook reference; this factory keeps request-level timeout and transport
''' option configuration available without weakening that boundary.
Public Function CreateRequest() As HttpRequest
    Set CreateRequest = New HttpRequest
End Function

''' Creates retry policy configuration for referenced-workbook consumers.
Public Function CreateRetryPolicy() As HttpRetryPolicy
    Set CreateRetryPolicy = New HttpRetryPolicy
End Function

''' Creates per-call reliability options for referenced-workbook consumers.
Public Function CreateExecutionOptions() As HttpExecutionOptions
    Set CreateExecutionOptions = New HttpExecutionOptions
End Function

''' Creates batch options for referenced-workbook consumers.
Public Function CreateBatchOptions() As HttpBatchOptions
    Set CreateBatchOptions = New HttpBatchOptions
End Function

''' Creates a cancellation token for referenced-workbook consumers.
Public Function CreateCancellationToken() As HttpCancellationToken
    Set CreateCancellationToken = New HttpCancellationToken
End Function

''' Creates native HTTP/2 and HTTP/3 negotiation options for consumers.
Public Function CreateProtocolOptions() As HttpProtocolOptions
    Set CreateProtocolOptions = New HttpProtocolOptions
End Function

''' Creates native response decompression options for consumers.
Public Function CreateDecompressionOptions() As HttpDecompressionOptions
    Set CreateDecompressionOptions = New HttpDecompressionOptions
End Function

''' Creates proxy routing options for referenced-workbook consumers.
Public Function CreateProxyOptions() As HttpProxyOptions
    Set CreateProxyOptions = New HttpProxyOptions
End Function

''' Creates an explicit caller-owned cookie jar for referenced-workbook consumers.
Public Function CreateCookieJar() As HttpCookieJar
    Set CreateCookieJar = New HttpCookieJar
End Function

''' Creates an opt-in structured diagnostics collector for referenced-workbook consumers.
Public Function CreateDiagnostics() As HttpDiagnostics
    Set CreateDiagnostics = New HttpDiagnostics
End Function

Public Function CreateBasicAuthProvider(ByVal Username As String, ByVal Password As String, Optional ByVal AllowInsecureHttp As Boolean = False) As IHttpAuthProvider
    Dim provider As New HttpBasicAuthProvider

    provider.Initialize Username, Password, AllowInsecureHttp
    Set CreateBasicAuthProvider = provider
End Function

Public Function CreateBearerAuthProvider(ByVal Token As String, Optional ByVal AllowInsecureHttp As Boolean = False) As IHttpAuthProvider
    Dim provider As New HttpBearerAuthProvider

    provider.Initialize Token, AllowInsecureHttp
    Set CreateBearerAuthProvider = provider
End Function

''' Creates a bounded WinHTTP challenge-auth provider for buffered requests.
Public Function CreateWindowsAuthProvider(ByVal Username As String, ByVal Password As String, Optional ByVal Scheme As HttpAuthChallengeScheme = HttpAuthSchemeAuto, Optional ByVal Target As HttpAuthChallengeTarget = HttpAuthTargetServer, Optional ByVal AllowInsecureHttp As Boolean = False, Optional ByVal MaxChallenges As Long = 3) As IHttpAuthProvider
    Dim provider As New HttpWindowsAuthProvider

    provider.Initialize Username, Password, Scheme, Target, AllowInsecureHttp, MaxChallenges
    Set CreateWindowsAuthProvider = provider
End Function

''' Creates an ordered multipart form for referenced-workbook consumers.
Public Function CreateMultipartForm() As HttpMultipartForm
    Set CreateMultipartForm = New HttpMultipartForm
End Function
