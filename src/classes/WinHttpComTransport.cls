VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpComTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHttpTransport

Private Const TransportSource As String = "WinHttpComTransport.Execute"
Private Const WinHttpOptionEnableRedirects As Long = 6
Private Const WinHttpOptionMaxAutomaticRedirects As Long = 14

Public Function Execute(ByVal Request As HttpRequest) As HttpResponse
    Set Execute = ExecuteCore(Request)
End Function

Private Function IHttpTransport_Execute(ByVal Request As HttpRequest) As HttpResponse
    Set IHttpTransport_Execute = ExecuteCore(Request)
End Function

Private Function ExecuteCore(ByVal Request As HttpRequest) As HttpResponse
    Dim backend As Object
    Dim started As Currency
    Dim finished As Currency
    Dim response As HttpResponse
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo TransportFailure
    If Request Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Request cannot be Nothing."
    End If

    Set backend = CreateObject("WinHttp.WinHttpRequest.5.1")
    backend.SetTimeouts Request.Timeouts.ResolveMilliseconds, Request.Timeouts.ConnectMilliseconds, Request.Timeouts.SendMilliseconds, Request.Timeouts.ReceiveMilliseconds
    backend.Option(WinHttpOptionEnableRedirects) = Request.FollowRedirects
    backend.Option(WinHttpOptionMaxAutomaticRedirects) = Request.MaxRedirects

    started = HttpTiming.CounterValue()
    backend.Open Request.Method, Request.Url, False
    ApplyRequestHeaders backend, Request.Headers
    SendRequest backend, Request.Body
    finished = HttpTiming.CounterValue()

    Set response = BuildResponse(backend, HttpTiming.ElapsedMilliseconds(started, finished))
    Set backend = Nothing
    Set ExecuteCore = response
    Exit Function

TransportFailure:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Set backend = Nothing

    If HttpErrors.CategoryFromNumber(errorNumber) <> HttpErrorNone Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If
    RaiseMappedFailure errorNumber
End Function

Private Sub ApplyRequestHeaders(ByVal backend As Object, ByVal Headers As HttpHeaders)
    Dim index As Long

    If backend Is Nothing Then
        Err.Raise HttpErrIo, TransportSource, "WinHTTP backend is unavailable."
        Exit Sub
    End If
    If Headers Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Request headers cannot be Nothing."
        Exit Sub
    End If
    For index = 1 To Headers.Count
        ' HttpHeaders validates field names and rejects CR/LF in values before this sink.
        backend.SetRequestHeader Headers.NameAt(index), Headers.ValueAt(index) ' xlflow:disable-line VBA224
    Next index
End Sub

Private Sub SendRequest(ByVal backend As Object, ByVal Body As HttpBody)
    Dim bytes As Variant

    If backend Is Nothing Then
        Err.Raise HttpErrIo, TransportSource, "WinHTTP backend is unavailable."
        Exit Sub
    End If
    If Body Is Nothing Then
        HttpErrors.RaiseValidation TransportSource, "Request body cannot be Nothing."
        Exit Sub
    End If
    If Body.IsEmpty Then
        backend.Send
    Else
        bytes = Body.Bytes
        backend.Send bytes
    End If
End Sub

Private Function BuildResponse(ByVal backend As Object, ByVal elapsedMilliseconds As Double) As HttpResponse
    Dim response As New HttpResponse
    Dim Headers As HttpHeaders
    Dim Body As New HttpBody
    Dim bytes As Variant

    If backend Is Nothing Then
        Err.Raise HttpErrIo, TransportSource, "WinHTTP backend is unavailable."
        Exit Function
    End If
    Set Headers = ParseResponseHeaders(CStr(backend.GetAllResponseHeaders()))
    bytes = backend.ResponseBody
    If HttpEncoding.ByteCount(bytes) > 0 Then Body.SetBytes bytes
    response.Initialize CLng(backend.Status), CStr(backend.StatusText), Headers, Body, elapsedMilliseconds
    Set BuildResponse = response
End Function

Private Function ParseResponseHeaders(ByVal rawHeaders As String) As HttpHeaders
    Dim Headers As New HttpHeaders
    Dim lines As Variant
    Dim line As Variant
    Dim separatorIndex As Long
    Dim text As String

    lines = Split(rawHeaders, vbCrLf)
    For Each line In lines
        text = CStr(line)
        If Len(text) > 0 Then
            separatorIndex = InStr(1, text, ":", vbBinaryCompare)
            If separatorIndex <= 1 Then
                Err.Raise HttpErrProtocol, TransportSource, "WinHTTP returned a malformed response header."
            End If
            Headers.Add Left$(text, separatorIndex - 1), Trim$(Mid$(text, separatorIndex + 1))
        End If
    Next line
    Set ParseResponseHeaders = Headers
End Function

Private Sub RaiseMappedFailure(ByVal nativeErrorNumber As Long)
    Dim code As Long
    Dim category As HttpErrorCategory
    Dim summary As String

    code = nativeErrorNumber And 65535
    category = CategoryFromWinHttpCode(code)
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
    HttpErrors.RaiseTransport category, TransportSource, summary & " (" & CStr(code) & ")."
End Sub

Private Function CategoryFromWinHttpCode(ByVal code As Long) As HttpErrorCategory
    Select Case code
    Case 12005, 12006
        CategoryFromWinHttpCode = HttpErrorInvalidUrl
    Case 12007
        CategoryFromWinHttpCode = HttpErrorDns
    Case 12029, 12030, 12031
        CategoryFromWinHttpCode = HttpErrorConnection
    Case 12037, 12038, 12044, 12045, 12055, 12057, 12157, 12175
        CategoryFromWinHttpCode = HttpErrorTls
    Case 12002
        CategoryFromWinHttpCode = HttpErrorTimeout
    Case 12017
        CategoryFromWinHttpCode = HttpErrorCancelled
    Case 12150 To 12156
        CategoryFromWinHttpCode = HttpErrorProtocol
    Case Else
        CategoryFromWinHttpCode = HttpErrorIo
    End Select
End Function
