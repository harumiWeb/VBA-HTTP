Attribute VB_Name = "HttpCookieParsing"
Option Explicit

Public Function ParseUrl(ByVal Url As String) As HttpCookieUrl
    Dim output As New HttpCookieUrl
    Dim schemeEnd As Long
    Dim authorityStart As Long
    Dim authorityEnd As Long
    Dim authority As String
    Dim path As String

    schemeEnd = InStr(1, Url, "://", vbBinaryCompare)
    If schemeEnd <= 1 Then HttpErrors.RaiseInvalidUrl "HttpCookieParsing.ParseUrl", "Cookie URL must be absolute HTTP or HTTPS."
    output.Scheme = Left$(Url, schemeEnd - 1)
    If output.Scheme <> "http" And output.Scheme <> "https" Then HttpErrors.RaiseInvalidUrl "HttpCookieParsing.ParseUrl", "Cookie URL must use HTTP or HTTPS."

    authorityStart = schemeEnd + 3
    authorityEnd = FirstUrlDelimiter(Url, authorityStart)
    authority = Mid$(Url, authorityStart, authorityEnd - authorityStart)
    authority = StripUserInfo(authority)
    authority = HostWithoutPort(authority)
    If Len(authority) = 0 Then HttpErrors.RaiseInvalidUrl "HttpCookieParsing.ParseUrl", "Cookie URL must contain a host."
    output.Host = authority

    path = "/"
    If authorityEnd <= Len(Url) And Mid$(Url, authorityEnd, 1) = "/" Then
        path = Mid$(Url, authorityEnd)
        path = BeforeDelimiter(path, "?")
        path = BeforeDelimiter(path, "#")
    End If
    If Len(path) = 0 Or Left$(path, 1) <> "/" Then path = "/"
    output.Path = path
    Set ParseUrl = output
End Function

Public Function TryParseSetCookie(ByVal HeaderValue As String, ByVal origin As HttpCookieUrl, ByVal nowUtc As Date) As HttpCookieRecord
    Dim segments As Variant
    Dim pairDelimiter As Long
    Dim attributeDelimiter As Long
    Dim pairName As String
    Dim pairValue As String
    Dim attributeName As String
    Dim attributeValue As String
    Dim index As Long
    Dim item As New HttpCookieRecord
    Dim maxAgeSeen As Boolean
    Dim maxAge As Double
    Dim expires As Date

    segments = Split(HeaderValue, ";")
    If Not IsArray(segments) Then Exit Function
    pairDelimiter = InStr(1, CStr(segments(0)), "=", vbBinaryCompare)
    If pairDelimiter <= 1 Then Exit Function
    pairName = Trim$(Left$(CStr(segments(0)), pairDelimiter - 1))
    pairValue = Mid$(CStr(segments(0)), pairDelimiter + 1)
    If Not IsCookieName(pairName) Or Not IsCookieValue(pairValue) Then Exit Function

    item.Name = pairName
    item.Value = pairValue
    item.Domain = origin.Host
    item.HostOnly = True
    item.Path = DefaultPath(origin.Path)

    If UBound(segments) < 1 Then
        Set TryParseSetCookie = item
        Exit Function
    End If
    For index = 1 To UBound(segments)
        attributeDelimiter = InStr(1, CStr(segments(index)), "=", vbBinaryCompare)
        If attributeDelimiter = 0 Then
            attributeName = Trim$(CStr(segments(index)))
            attributeValue = ""
        Else
            attributeName = Trim$(Left$(CStr(segments(index)), attributeDelimiter - 1))
            attributeValue = Mid$(CStr(segments(index)), attributeDelimiter + 1)
        End If
        Select Case LCase$(attributeName)
        Case "domain"
            If attributeDelimiter = 0 Then Exit Function
            item.Domain = LCase$(Trim$(attributeValue))
            If Left$(item.Domain, 1) = "." Then item.Domain = Mid$(item.Domain, 2)
            If Len(item.Domain) = 0 Or Not DomainMatches(origin.Host, item.Domain) Then Exit Function
            item.HostOnly = False
        Case "path"
            If attributeDelimiter = 0 Then Exit Function
            If Len(attributeValue) = 0 Or Left$(attributeValue, 1) <> "/" Then Exit Function
            item.Path = attributeValue
        Case "secure"
            item.Secure = True
        Case "max-age"
            If attributeDelimiter = 0 Then Exit Function
            If Not TryParseDouble(attributeValue, maxAge) Then Exit Function
            maxAgeSeen = True
        Case "expires"
            If attributeDelimiter > 0 Then
                If HttpCookieDates.TryParse(attributeValue, expires) Then item.ExpiresAt = expires: item.HasExpires = True
            End If
        End Select
    Next index

    If maxAgeSeen Then
        item.HasExpires = True
        If maxAge <= 0 Then
            item.ExpiresAt = nowUtc
        Else
            On Error GoTo InvalidExpiry
            item.ExpiresAt = DateAdd("s", maxAge, nowUtc)
        End If
    End If
    Set TryParseSetCookie = item
    Exit Function
InvalidExpiry:
    Set TryParseSetCookie = Nothing
    Err.Clear
End Function

Public Function DomainMatches(ByVal host As String, ByVal domain As String) As Boolean
    host = LCase$(host)
    domain = LCase$(domain)
    If host = domain Then
        DomainMatches = True
    ElseIf Len(host) > Len(domain) And Right$(host, Len(domain) + 1) = "." & domain Then
        DomainMatches = True
    End If
End Function

Public Function PathMatches(ByVal requestPath As String, ByVal cookiePath As String) As Boolean
    If requestPath = cookiePath Then
        PathMatches = True
    ElseIf Left$(requestPath, Len(cookiePath)) = cookiePath Then
        PathMatches = (Right$(cookiePath, 1) = "/" Or Mid$(requestPath, Len(cookiePath) + 1, 1) = "/")
    End If
End Function

Public Function DefaultPath(ByVal requestPath As String) As String
    Dim slash As Long
    If Len(requestPath) = 0 Or Left$(requestPath, 1) <> "/" Then DefaultPath = "/": Exit Function
    slash = InStrRev(requestPath, "/")
    If slash <= 1 Then DefaultPath = "/" Else DefaultPath = Left$(requestPath, slash - 1)
End Function

Private Function FirstUrlDelimiter(ByVal value As String, ByVal startAt As Long) As Long
    Dim index As Long
    For index = startAt To Len(value)
        If InStr(1, "/?#", Mid$(value, index, 1), vbBinaryCompare) > 0 Then FirstUrlDelimiter = index: Exit Function
    Next index
    FirstUrlDelimiter = Len(value) + 1
End Function

Private Function StripUserInfo(ByVal authority As String) As String
    If InStrRev(authority, "@") > 0 Then authority = Mid$(authority, InStrRev(authority, "@") + 1)
    StripUserInfo = authority
End Function

Private Function HostWithoutPort(ByVal authority As String) As String
    Dim closeBracket As Long
    If Left$(authority, 1) = "[" Then
        closeBracket = InStr(1, authority, "]", vbBinaryCompare)
        If closeBracket > 0 Then HostWithoutPort = Mid$(authority, 2, closeBracket - 2): Exit Function
    End If
    If InStr(1, authority, ":", vbBinaryCompare) > 0 Then
        HostWithoutPort = Left$(authority, InStrRev(authority, ":") - 1)
    Else
        HostWithoutPort = authority
    End If
End Function

Private Function BeforeDelimiter(ByVal value As String, ByVal delimiter As String) As String
    Dim at As Long
    at = InStr(1, value, delimiter, vbBinaryCompare)
    If at = 0 Then BeforeDelimiter = value Else BeforeDelimiter = Left$(value, at - 1)
End Function

Private Function IsCookieName(ByVal value As String) As Boolean
    Dim index As Long
    Dim code As Long
    Dim character As String
    If Len(value) = 0 Then Exit Function
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        code = AscW(character)
        If code < 0 Then code = code + 65536
        If code <= 32 Or code >= 127 Or InStr(1, "()<>@,;:\" & Chr$(34) & "/[]?={} ", character, vbBinaryCompare) > 0 Then Exit Function
    Next index
    IsCookieName = True
End Function

Private Function IsCookieValue(ByVal value As String) As Boolean
    Dim index As Long
    Dim code As Long
    For index = 1 To Len(value)
        code = AscW(Mid$(value, index, 1))
        If code < 32 Or code = 127 Or code = 59 Then Exit Function
    Next index
    IsCookieValue = True
End Function

Private Function TryParseDouble(ByVal value As String, ByRef output As Double) As Boolean
    On Error GoTo Invalid
    If Len(Trim$(value)) = 0 Then Exit Function
    output = CDbl(Trim$(value))
    TryParseDouble = True
    Exit Function
Invalid:
    TryParseDouble = False
    Err.Clear
End Function
