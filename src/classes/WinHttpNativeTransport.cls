VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpNativeTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHttpTransport

Private Const TransportSource As String = "WinHttpNativeTransport.Execute"
Private Const UserAgent As String = "VBA-HTTP/0.5"
Private Const ReadChunkBytes As Long = 65536

Public Function Execute(ByVal Request As HttpRequest) As HttpResponse
    Set Execute = ExecuteCore(Request)
End Function

Private Function IHttpTransport_Execute(ByVal Request As HttpRequest) As HttpResponse
    Set IHttpTransport_Execute = ExecuteCore(Request)
End Function

Private Function ExecuteCore(ByVal Request As HttpRequest) As HttpResponse
    Dim Url As WinHttpNativeUrl
    Dim session As WinHttpSessionHandle
    Dim connection As WinHttpConnectionHandle
    Dim nativeRequest As WinHttpRequestHandle
    Dim response As HttpResponse
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    #If VBA7 Then
    Dim handleValue As LongPtr
    #Else
    Dim handleValue As Long
    #End If
    Dim started As Currency

    On Error GoTo NativeFailure
    If Request Is Nothing Then HttpErrors.RaiseValidation TransportSource, "Request cannot be Nothing."

    started = HttpTiming.CounterValue()
    Set Url = New WinHttpNativeUrl
    Url.Initialize Request.Url
    Set session = New WinHttpSessionHandle
    handleValue = WinHttpNativeApi.OpenSession(UserAgent)
    If handleValue = 0 Then RaiseNativeFailure "OpenSession"
    session.Initialize handleValue

    Set connection = New WinHttpConnectionHandle
    handleValue = WinHttpNativeApi.Connect(session.Value, Url.Host, Url.Port)
    If handleValue = 0 Then RaiseNativeFailure "Connect"
    connection.Initialize handleValue, session

    Set nativeRequest = New WinHttpRequestHandle
    If Url.IsSecure Then
        handleValue = WinHttpNativeApi.OpenRequest(connection.Value, Request.Method, Url.PathAndQuery, WinHttpNativeApi.WinHttpFlagSecure)
    Else
        handleValue = WinHttpNativeApi.OpenRequest(connection.Value, Request.Method, Url.PathAndQuery, 0)
    End If
    If handleValue = 0 Then RaiseNativeFailure "OpenRequest"
    nativeRequest.Initialize handleValue, connection

    ApplyRequestOptions nativeRequest, Request
    If Not WinHttpNativeApi.AddRequestHeaders(nativeRequest.Value, BuildRequestHeaders(Request.Headers)) Then RaiseNativeFailure "AddRequestHeaders"
    If Not WinHttpNativeApi.SendBody(nativeRequest.Value, Request.Body.Bytes) Then RaiseNativeFailure "SendRequest"
    If Not WinHttpNativeApi.ReceiveResponse(nativeRequest.Value) Then RaiseNativeFailure "ReceiveResponse"

    Set response = BuildResponse(nativeRequest, HttpTiming.ElapsedMilliseconds(started, HttpTiming.CounterValue()))
    Set ExecuteCore = response
    CloseNativeHandles nativeRequest, connection, session
    Exit Function

NativeFailure:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    CloseNativeHandles nativeRequest, connection, session

    If errorNumber <> 0 Then
        If HttpErrors.CategoryFromNumber(errorNumber) <> HttpErrorNone Then
            Err.Raise errorNumber, errorSource, errorDescription
        Else
            WinHttpErrorMapping.RaiseMappedWinHttpFailure errorNumber, TransportSource
        End If
        Err.Raise errorNumber, errorSource, errorDescription
    Else
        Err.Raise HttpErrIo, TransportSource, "Native WinHTTP failure did not provide an error number."
    End If
End Function

Private Sub CloseNativeHandles(ByVal nativeRequest As WinHttpRequestHandle, ByVal connection As WinHttpConnectionHandle, ByVal session As WinHttpSessionHandle)
    If Not nativeRequest Is Nothing Then nativeRequest.ReleaseHandle
    If Not connection Is Nothing Then connection.ReleaseHandle
    If Not session Is Nothing Then session.ReleaseHandle
End Sub

Private Sub RaiseNativeFailure(Optional ByVal Operation As String = "")
    WinHttpErrorMapping.RaiseMappedWinHttpFailure WinHttpNativeApi.LastErrorCode(), TransportSource, Operation
End Sub

Private Sub ApplyRequestOptions(ByVal nativeRequest As WinHttpRequestHandle, ByVal Request As HttpRequest)
    Dim redirectFlag As Long

    If Not WinHttpNativeApi.SetTimeouts(nativeRequest.Value, Request.Timeouts.ResolveMilliseconds, Request.Timeouts.ConnectMilliseconds, Request.Timeouts.SendMilliseconds, Request.Timeouts.ReceiveMilliseconds) Then RaiseNativeFailure "SetTimeouts"
    If Request.FollowRedirects Then
        redirectFlag = 1
    Else
        redirectFlag = 0
    End If
    If Not WinHttpNativeApi.SetOptionLong(nativeRequest.Value, WinHttpNativeApi.WinHttpOptionRedirectPolicy, redirectFlag) Then RaiseNativeFailure "SetRedirects"
    If Not WinHttpNativeApi.SetOptionLong(nativeRequest.Value, WinHttpNativeApi.WinHttpOptionMaxAutomaticRedirects, Request.MaxRedirects) Then RaiseNativeFailure "SetMaxRedirects"
End Sub

Private Function BuildRequestHeaders(ByVal Headers As HttpHeaders) As String
    Dim index As Long
    Dim result As String

    If Headers Is Nothing Then HttpErrors.RaiseValidation TransportSource, "Request headers cannot be Nothing."
    For index = 1 To Headers.Count
        result = result & Headers.NameAt(index) & ": " & Headers.ValueAt(index) & vbCrLf
    Next index
    BuildRequestHeaders = result
End Function

Private Function BuildResponse(ByVal nativeRequest As WinHttpRequestHandle, ByVal elapsedMilliseconds As Double) As HttpResponse
    Dim statusText As String
    Dim rawHeaders As String
    Dim statusCode As Long
    Dim Body As HttpBody
    Dim protocolUsed As String
    Dim response As New HttpResponse

    If Not WinHttpNativeApi.QueryHeaderString(nativeRequest.Value, WinHttpNativeApi.WinHttpQueryStatusCode, statusText) Then RaiseNativeFailure "QueryStatusCode"
    On Error GoTo InvalidStatus
    statusCode = CLng(Trim$(statusText))
    On Error GoTo 0
    If statusCode < 100 Or statusCode > 599 Then GoTo InvalidStatus
    If Not WinHttpNativeApi.QueryHeaderString(nativeRequest.Value, WinHttpNativeApi.WinHttpQueryStatusText, statusText) Then RaiseNativeFailure "QueryStatusText"
    If Not WinHttpNativeApi.QueryHeaderString(nativeRequest.Value, WinHttpNativeApi.WinHttpQueryRawHeadersCrlf, rawHeaders) Then RaiseNativeFailure "QueryRawHeaders"
    Set Body = ReadBufferedBody(nativeRequest)
    protocolUsed = QueryProtocol(nativeRequest)

    response.Initialize statusCode, statusText, ParseResponseHeaders(rawHeaders), Body, elapsedMilliseconds, protocolUsed
    Set BuildResponse = response
    Exit Function

InvalidStatus:
    On Error GoTo 0
    Err.Raise HttpErrProtocol, TransportSource, "WinHTTP returned an invalid response status."
End Function

Private Function ReadBufferedBody(ByVal nativeRequest As WinHttpRequestHandle) As HttpBody
    Dim Body As New HttpBody
    Dim output() As Byte
    Dim chunk() As Byte
    Dim capacity As Long
    Dim total As Long
    Dim available As Long
    Dim requested As Long
    Dim readBytes As Long

    Do
        If Not WinHttpNativeApi.QueryDataAvailable(nativeRequest.Value, available) Then RaiseNativeFailure "QueryDataAvailable"
        If available = 0 Then Exit Do
        requested = available
        If requested > ReadChunkBytes Then requested = ReadChunkBytes
        ReDim chunk(0 To requested - 1)
        If Not WinHttpNativeApi.ReadData(nativeRequest.Value, chunk, requested, readBytes) Then RaiseNativeFailure "ReadData"
        If readBytes = 0 Then Exit Do
        If total > 2147483647 - readBytes Then HttpErrors.RaiseTransport HttpErrorIo, TransportSource, "Native buffered response exceeds the VBA byte-array limit."
        EnsureCapacity output, capacity, total + readBytes
        WinHttpNativeApi.CopyByteRange output, total, chunk, readBytes
        total = total + readBytes
    Loop

    If total > 0 Then
        ReDim Preserve output(0 To total - 1)
        Body.SetBytes output
    End If
    Set ReadBufferedBody = Body
End Function

Private Sub EnsureCapacity(ByRef buffer() As Byte, ByRef capacity As Long, ByVal required As Long)
    Dim nextCapacity As Long

    If required <= capacity Then Exit Sub
    If capacity = 0 Then
        nextCapacity = ReadChunkBytes
    Else
        nextCapacity = capacity
    End If
    Do While nextCapacity < required
        If nextCapacity > 1073741823 Then
            nextCapacity = required
            Exit Do
        End If
        nextCapacity = nextCapacity * 2
    Loop
    If capacity = 0 Then
        ReDim buffer(0 To nextCapacity - 1)
    Else
        ReDim Preserve buffer(0 To nextCapacity - 1)
    End If
    capacity = nextCapacity
End Sub

Private Function ParseResponseHeaders(ByVal rawHeaders As String) As HttpHeaders
    Dim Headers As New HttpHeaders
    Dim lines As Variant
    Dim line As Variant
    Dim text As String
    Dim separatorIndex As Long

    lines = Split(rawHeaders, vbCrLf)
    For Each line In lines
        text = CStr(line)
        If Len(text) > 0 Then
            If Left$(text, 5) = "HTTP/" Then GoTo NextHeader
            separatorIndex = InStr(1, text, ":", vbBinaryCompare)
            If separatorIndex <= 1 Then Err.Raise HttpErrProtocol, TransportSource, "WinHTTP returned a malformed response header."
            Headers.Add Left$(text, separatorIndex - 1), Trim$(Mid$(text, separatorIndex + 1))
        End If
NextHeader:
    Next line
    Set ParseResponseHeaders = Headers
End Function

Private Function QueryProtocol(ByVal nativeRequest As WinHttpRequestHandle) As String
    Dim protocolFlags As Long

    If Not WinHttpNativeApi.QueryProtocolUsed(nativeRequest.Value, protocolFlags) Then
        If WinHttpNativeApi.LastErrorCode() <> WinHttpNativeApi.WinHttpErrorInvalidOption Then RaiseNativeFailure "QueryProtocolUsed"
    End If
    If (protocolFlags And WinHttpNativeApi.WinHttpProtocolFlagHttp3) <> 0 Then
        QueryProtocol = "HTTP/3"
    ElseIf (protocolFlags And WinHttpNativeApi.WinHttpProtocolFlagHttp2) <> 0 Then
        QueryProtocol = "HTTP/2"
    Else
        QueryProtocol = "HTTP/1.1"
    End If
End Function
