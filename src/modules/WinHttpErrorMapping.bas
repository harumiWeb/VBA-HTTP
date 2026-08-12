Attribute VB_Name = "WinHttpErrorMapping"
Option Explicit

Public Function CategoryFromCode(ByVal code As Long) As HttpErrorCategory
    Select Case code
    Case 12005, 12006
        CategoryFromCode = HttpErrorInvalidUrl
    Case 12007
        CategoryFromCode = HttpErrorDns
    Case 12029, 12030, 12031
        CategoryFromCode = HttpErrorConnection
    Case 12037, 12038, 12044, 12045, 12055, 12057, 12157, 12175
        CategoryFromCode = HttpErrorTls
    Case 12002
        CategoryFromCode = HttpErrorTimeout
    Case 12017
        CategoryFromCode = HttpErrorCancelled
    Case 12150, 12152, 12154, 12155, 12156
        CategoryFromCode = HttpErrorProtocol
    Case Else
        CategoryFromCode = HttpErrorIo
    End Select
End Function

Public Sub RaiseMappedWinHttpFailure(ByVal nativeErrorNumber As Long, ByVal Source As String, Optional ByVal Operation As String = "")
    Dim code As Long
    Dim category As HttpErrorCategory
    Dim summary As String

    code = nativeErrorNumber And 65535
    category = CategoryFromCode(code)
    Select Case category
    Case HttpErrorInvalidUrl
        summary = "WinHTTP rejected the request URL"
    Case HttpErrorDns
        summary = "WinHTTP could not resolve the server name"
    Case HttpErrorConnection
        summary = "WinHTTP could not establish or maintain the connection"
    Case HttpErrorTls
        summary = "WinHTTP TLS validation failed"
    Case HttpErrorTimeout
        summary = "WinHTTP request timed out"
    Case HttpErrorCancelled
        summary = "WinHTTP request was cancelled"
    Case HttpErrorProtocol
        summary = "WinHTTP reported an invalid HTTP exchange"
    Case Else
        category = HttpErrorIo
        summary = "WinHTTP request failed"
    End Select
    If Len(Operation) > 0 Then summary = summary & " during " & Operation
    HttpErrors.RaiseTransport category, Source, summary & " (" & CStr(code) & ")."
End Sub
