Attribute VB_Name = "HttpCookieJarTests"
Option Explicit

Public Sub Test_CookieJar_StoresAndMatchesHostOnlyCookie()
    Dim jar As New HttpCookieJar
    Dim headers As New HttpHeaders

    headers.Add "Set-Cookie", "session=alpha; Path=/cookie"
    jar.StoreResponseCookies "http://example.test/cookie/set", headers

    XlflowAssert.AssertEquals 1, jar.Count
    XlflowAssert.AssertEquals "session=alpha", jar.GetCookieHeader("http://example.test/cookie/echo")
    XlflowAssert.AssertEquals "", jar.GetCookieHeader("http://other.example.test/cookie/echo")
End Sub

Public Sub Test_CookieJar_OrdersLongestPathFirstAndHonorsSecure()
    Dim jar As New HttpCookieJar
    Dim headers As New HttpHeaders

    headers.Add "Set-Cookie", "root=one; Path=/"
    headers.Add "Set-Cookie", "nested=two; Path=/cookie"
    headers.Add "Set-Cookie", "secure=three; Path=/; Secure"
    jar.StoreResponseCookies "https://example.test/cookie/set", headers

    XlflowAssert.AssertEquals "nested=two; root=one; secure=three", jar.GetCookieHeader("https://example.test/cookie/echo")
    XlflowAssert.AssertEquals "nested=two; root=one", jar.GetCookieHeader("http://example.test/cookie/echo")
End Sub

Public Sub Test_CookieJar_DeletesWithMaxAgeZero()
    Dim jar As New HttpCookieJar
    Dim headers As New HttpHeaders

    headers.Add "Set-Cookie", "session=alpha; Path=/cookie"
    jar.StoreResponseCookies "http://example.test/cookie/set", headers
    headers.Clear
    headers.Add "Set-Cookie", "session=; Max-Age=0; Path=/cookie"
    jar.StoreResponseCookies "http://example.test/cookie/clear", headers

    XlflowAssert.AssertEquals 0, jar.Count
    XlflowAssert.AssertEquals "", jar.GetCookieHeader("http://example.test/cookie/echo")
End Sub

Public Sub Test_CookieJar_CallerCookieHeaderRemainsAuthoritative()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim jar As New HttpCookieJar
    Dim cookieResponseHeaders As New HttpHeaders
    Dim Request As New HttpRequest

    cookieResponseHeaders.Add "Set-Cookie", "session=server; Path=/"
    configuredResponse.Initialize 204, "", cookieResponseHeaders
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    Set client.CookieJar = jar
    jar.StoreResponseCookies "https://example.test/set", cookieResponseHeaders

    Request.Url = "https://example.test/echo"
    Request.Headers.SetValue "Cookie", "caller=authoritative"
    Call client.Execute(Request)

    XlflowAssert.AssertEquals "caller=authoritative", transport.LastRequest.Headers.GetValue("Cookie")
    XlflowAssert.AssertFalse transport.LastRequest.FollowRedirects
    XlflowAssert.AssertTrue Request.FollowRedirects
End Sub

Public Sub Test_CookieJar_ClientAndRequestSnapshotPrecedence()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim clientJar As New HttpCookieJar
    Dim requestJar As New HttpCookieJar
    Dim clientHeaders As New HttpHeaders
    Dim requestHeaders As New HttpHeaders
    Dim Request As New HttpRequest
    Dim clientRequest As New HttpRequest

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    Set client.CookieJar = clientJar
    clientHeaders.Add "Set-Cookie", "client=one; Path=/"
    clientJar.StoreResponseCookies "https://example.test/set", clientHeaders
    requestHeaders.Add "Set-Cookie", "request=two; Path=/"
    requestJar.StoreResponseCookies "https://example.test/set", requestHeaders

    Request.Url = "https://example.test/echo"
    Set Request.CookieJar = requestJar
    Call client.Execute(Request)

    XlflowAssert.AssertEquals "request=two", transport.LastRequest.Headers.GetValue("Cookie")
    XlflowAssert.AssertSame requestJar, transport.LastRequest.CookieJar
    XlflowAssert.AssertSame clientJar, client.CookieJar

    clientRequest.Url = "https://example.test/echo"
    Call client.Execute(clientRequest)
    XlflowAssert.AssertEquals "client=one", transport.LastRequest.Headers.GetValue("Cookie")
    XlflowAssert.AssertSame clientJar, transport.LastRequest.CookieJar
End Sub

Public Sub Test_CookieJar_ExpiresAgainstInjectedClock()
    Dim jar As New HttpCookieJar
    Dim clock As New FakeCookieClock
    Dim headers As New HttpHeaders

    clock.NowUtc = DateSerial(2026, 1, 1)
    Set jar.Clock = clock
    headers.Add "Set-Cookie", "short=one; Max-Age=10; Path=/"
    jar.StoreResponseCookies "http://example.test/set", headers
    XlflowAssert.AssertEquals "short=one", jar.GetCookieHeader("http://example.test/")

    clock.NowUtc = DateSerial(2026, 1, 1) + TimeSerial(0, 0, 11)
    XlflowAssert.AssertEquals "", jar.GetCookieHeader("http://example.test/")
End Sub
