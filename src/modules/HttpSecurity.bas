Attribute VB_Name = "HttpSecurity"
Option Explicit

Public Function IsSensitiveHeader(ByVal Name As String) As Boolean
    Select Case LCase$(Trim$(Name))
    Case "authorization", "proxy-authorization", "cookie", "set-cookie", "www-authenticate", "proxy-authenticate"
        IsSensitiveHeader = True
    End Select
End Function

Public Function IsRedirectSensitiveHeader(ByVal Name As String) As Boolean
    Select Case LCase$(Trim$(Name))
    Case "authorization", "proxy-authorization", "cookie"
        IsRedirectSensitiveHeader = True
    End Select
End Function

Public Function HasRedirectSensitiveHeaders(ByVal Headers As HttpHeaders) As Boolean
    Dim index As Long

    If Headers Is Nothing Then Exit Function
    For index = 1 To Headers.Count
        If IsRedirectSensitiveHeader(Headers.NameAt(index)) Then
            HasRedirectSensitiveHeaders = True
            Exit Function
        End If
    Next index
End Function

Public Function RedactHeaderValue(ByVal Name As String, ByVal value As String) As String
    If IsSensitiveHeader(Name) Then
        RedactHeaderValue = "[REDACTED]"
    Else
        RedactHeaderValue = value
    End If
End Function

Public Function RedactedHeaders(ByVal Headers As HttpHeaders) As HttpHeaders
    Dim output As New HttpHeaders
    Dim index As Long

    If Headers Is Nothing Then
        Set RedactedHeaders = output
        Exit Function
    End If
    For index = 1 To Headers.Count
        output.Add Headers.NameAt(index), RedactHeaderValue(Headers.NameAt(index), Headers.ValueAt(index))
    Next index
    Set RedactedHeaders = output
End Function
