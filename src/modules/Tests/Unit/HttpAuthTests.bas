Attribute VB_Name = "HttpAuthTests"
Option Explicit

Public Sub Test_Auth_BasicEncodesAuthorizationAndDisablesRedirects()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim provider As IHttpAuthProvider

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    client.BaseUrl = "https://example.test"
    Set provider = VBAHttp.CreateBasicAuthProvider("user", "pass")
    Set client.AuthProvider = provider

    Call client.GetResponse("/protected")

    XlflowAssert.AssertEquals "Basic dXNlcjpwYXNz", transport.LastRequest.Headers.GetValue("Authorization")
    XlflowAssert.AssertFalse transport.LastRequest.FollowRedirects
    XlflowAssert.AssertNotSame provider, transport.LastRequest.AuthProvider
End Sub

Public Sub Test_Auth_BasicUsesUtf8BeforeBase64()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim provider As IHttpAuthProvider
    Dim username As String
    Dim password As String
    Dim expectedHeader As String

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    client.BaseUrl = "https://example.test"
    username = "ユーザー"
    password = "p" & ChrW$(228) & "ss"
    Set provider = VBAHttp.CreateBasicAuthProvider(username, password)
    Set client.AuthProvider = provider

    Call client.GetResponse("/protected")

    expectedHeader = "Basic " & HttpEncoding.EncodeBase64(HttpEncoding.EncodeUtf8(username & ":" & password))
    XlflowAssert.AssertEquals expectedHeader, transport.LastRequest.Headers.GetValue("Authorization")
End Sub

Public Sub Test_Auth_Base64EncodesBinaryTriples()
    Dim bytes(0 To 2) As Byte

    bytes(0) = 0
    bytes(1) = 255
    bytes(2) = 254
    XlflowAssert.AssertEquals "AP/+", HttpEncoding.EncodeBase64(bytes)
End Sub

Public Sub Test_Auth_BearerSetsOpaqueTokenAndDisablesRedirects()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim provider As IHttpAuthProvider

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    Set provider = VBAHttp.CreateBearerAuthProvider("vba-http-token", True)
    Set client.AuthProvider = provider

    Call client.GetResponse("http://example.test/protected")

    XlflowAssert.AssertEquals "Bearer vba-http-token", transport.LastRequest.Headers.GetValue("Authorization")
    XlflowAssert.AssertFalse transport.LastRequest.FollowRedirects
End Sub

Public Sub Test_Auth_RequestProviderOverridesClientProvider()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim clientProvider As IHttpAuthProvider
    Dim requestProvider As IHttpAuthProvider
    Dim Request As New HttpRequest

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    Set clientProvider = VBAHttp.CreateBearerAuthProvider("client-token", True)
    Set requestProvider = VBAHttp.CreateBearerAuthProvider("request-token", True)
    Set client.AuthProvider = clientProvider
    Request.Url = "http://example.test/protected"
    Set Request.AuthProvider = requestProvider

    Call client.Execute(Request)

    XlflowAssert.AssertEquals "Bearer request-token", transport.LastRequest.Headers.GetValue("Authorization")
    XlflowAssert.AssertEquals "", Request.Headers.GetValue("Authorization")
End Sub

'@ExpectedError(-2147200503, "Authorization header conflicts with the configured auth provider.", "HttpBasicAuthProvider.Apply")
Public Sub Test_Auth_RejectsExistingAuthorizationHeader()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim provider As IHttpAuthProvider

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    client.DefaultHeaders.SetValue "Authorization", "Bearer caller-value"
    Set provider = VBAHttp.CreateBasicAuthProvider("user", "pass")
    Set client.AuthProvider = provider

    Call client.GetResponse("https://example.test/protected")
End Sub

'@ExpectedError(-2147200503, "Basic authentication requires an HTTPS request URL.", "HttpBasicAuthProvider.Apply")
Public Sub Test_Auth_BasicRequiresHttpsByDefault()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim provider As IHttpAuthProvider

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    Set provider = VBAHttp.CreateBasicAuthProvider("user", "pass")
    Set client.AuthProvider = provider

    Call client.GetResponse("http://example.test/protected")
End Sub

'@ExpectedError(-2147200503, "Basic authentication requires an HTTP or HTTPS request URL.", "HttpBasicAuthProvider.Apply")
Public Sub Test_Auth_RejectsNonHttpUrlEvenWhenInsecureIsAllowed()
    Dim provider As IHttpAuthProvider
    Dim Request As New HttpRequest

    Set provider = VBAHttp.CreateBasicAuthProvider("user", "pass", True)
    Request.Url = "ftp://example.test/protected"
    provider.Apply Request
End Sub

'@ExpectedError(-2147200503, "Basic username cannot contain a colon.", "HttpBasicAuthProvider.Initialize")
Public Sub Test_Auth_BasicRejectsColonInUsername()
    Call VBAHttp.CreateBasicAuthProvider("user:name", "pass")
End Sub

'@ExpectedError(-2147200503, "Bearer token cannot contain whitespace or control characters.", "HttpBearerAuthProvider.Initialize")
Public Sub Test_Auth_BearerRejectsWhitespace()
    Call VBAHttp.CreateBearerAuthProvider("token with-space")
End Sub

'@ExpectedError(-2147200503, "Bearer token must contain ASCII characters only.", "HttpBearerAuthProvider.Initialize")
Public Sub Test_Auth_BearerRejectsNonAscii()
    Call VBAHttp.CreateBearerAuthProvider("トークン")
End Sub

Public Sub Test_Security_RedactsSensitiveHeaderValues()
    XlflowAssert.AssertEquals "[REDACTED]", HttpSecurity.RedactHeaderValue("Authorization", "secret")
    XlflowAssert.AssertEquals "[REDACTED]", HttpSecurity.RedactHeaderValue("proxy-authorization", "secret")
    XlflowAssert.AssertEquals "safe", HttpSecurity.RedactHeaderValue("X-Trace", "safe")
    XlflowAssert.AssertTrue HttpSecurity.IsSensitiveHeader("WWW-Authenticate")
End Sub
