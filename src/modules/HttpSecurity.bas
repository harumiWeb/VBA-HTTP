Attribute VB_Name = "HttpSecurity"
Option Explicit

Public Function IsSensitiveHeader(ByVal Name As String) As Boolean
    Select Case LCase$(Trim$(Name))
    Case "authorization", "proxy-authorization", "cookie", "set-cookie", "www-authenticate", "proxy-authenticate"
        IsSensitiveHeader = True
    End Select
End Function

Public Function RedactHeaderValue(ByVal Name As String, ByVal value As String) As String
    If IsSensitiveHeader(Name) Then
        RedactHeaderValue = "[REDACTED]"
    Else
        RedactHeaderValue = value
    End If
End Function
