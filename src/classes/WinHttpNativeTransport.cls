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
Implements IHttpDownloadTransport
Implements IHttpUploadTransport

Private Const TransportSource As String = "WinHttpNativeTransport.Execute"
Private Const DownloadSource As String = "WinHttpNativeTransport.DownloadFile"
Private Const UploadSource As String = "WinHttpNativeTransport.Upload"
Private Const UserAgent As String = "VBA-HTTP/0.5"
Private Const ReadChunkBytes As Long = 65536

Public Function Execute(ByVal Request As HttpRequest) As HttpResponse
    Set Execute = ExecuteCore(Request)
End Function

Private Function IHttpDownloadTransport_DownloadFile(ByVal Request As HttpRequest, ByVal DestinationPath As String, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing) As HttpDownloadResult
    Set IHttpDownloadTransport_DownloadFile = DownloadCore(Request, DestinationPath, Options, Progress)
End Function

Private Function IHttpUploadTransport_UploadFile(ByVal Request As HttpRequest, ByVal SourcePath As String, ByVal ContentType As String, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing) As HttpUploadResult
    Set IHttpUploadTransport_UploadFile = UploadFileCore(Request, SourcePath, ContentType, Options, Progress)
End Function

Private Function IHttpUploadTransport_UploadMultipart(ByVal Request As HttpRequest, ByVal Form As HttpMultipartForm, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing) As HttpUploadResult
    Set IHttpUploadTransport_UploadMultipart = UploadMultipartCore(Request, Form, Options, Progress)
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
    handleValue = OpenSessionForRequest(Request, TransportSource)
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

Private Function DownloadCore(ByVal Request As HttpRequest, ByVal DestinationPath As String, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing) As HttpDownloadResult
    Dim Url As WinHttpNativeUrl
    Dim session As WinHttpSessionHandle
    Dim connection As WinHttpConnectionHandle
    Dim nativeRequest As WinHttpRequestHandle
    Dim writer As WinHttpDownloadFileWriter
    Dim result As HttpDownloadResult
    Dim responseHeaders As HttpHeaders
    Dim destination As String
    Dim temporaryPath As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim statusCode As Long
    Dim reasonPhrase As String
    Dim protocolUsed As String
    Dim contentLength As Currency
    Dim contentLengthKnown As Boolean
    Dim bytesWritten As Currency
    Dim started As Currency
    Dim available As Long
    Dim requested As Long
    Dim readBytes As Long
    Dim chunk() As Byte
    Dim finalChunk() As Byte
    Dim copyIndex As Long
    Dim cleanupSucceeded As Boolean
    #If VBA7 Then
    Dim handleValue As LongPtr
    #Else
    Dim handleValue As Long
    #End If

    On Error GoTo DownloadFailure
    If Request Is Nothing Then HttpErrors.RaiseValidation DownloadSource, "Request cannot be Nothing."
    If Request.Method <> "GET" Then HttpErrors.RaiseValidation DownloadSource, "Streaming downloads require the GET method."
    If Options Is Nothing Then HttpErrors.RaiseValidation DownloadSource, "Execution options cannot be Nothing."

    destination = WinHttpDownloadFileSystem.NormalizeDestination(DestinationPath)
    CheckDownloadControls Options, 0
    started = HttpTiming.CounterValue()
    Set Url = New WinHttpNativeUrl
    Url.Initialize Request.Url
    Set session = New WinHttpSessionHandle
    handleValue = OpenSessionForRequest(Request, DownloadSource)
    If handleValue = 0 Then RaiseNativeFailure "OpenSession", DownloadSource
    session.Initialize handleValue

    Set connection = New WinHttpConnectionHandle
    handleValue = WinHttpNativeApi.Connect(session.Value, Url.Host, Url.Port)
    If handleValue = 0 Then RaiseNativeFailure "Connect", DownloadSource
    connection.Initialize handleValue, session

    Set nativeRequest = New WinHttpRequestHandle
    If Url.IsSecure Then
        handleValue = WinHttpNativeApi.OpenRequest(connection.Value, Request.Method, Url.PathAndQuery, WinHttpNativeApi.WinHttpFlagSecure)
    Else
        handleValue = WinHttpNativeApi.OpenRequest(connection.Value, Request.Method, Url.PathAndQuery, 0)
    End If
    If handleValue = 0 Then RaiseNativeFailure "OpenRequest", DownloadSource
    nativeRequest.Initialize handleValue, connection

    ApplyRequestOptions nativeRequest, Request, DownloadSource
    If Not WinHttpNativeApi.AddRequestHeaders(nativeRequest.Value, BuildRequestHeaders(Request.Headers)) Then RaiseNativeFailure "AddRequestHeaders", DownloadSource
    If Not WinHttpNativeApi.SendBody(nativeRequest.Value, Request.Body.Bytes) Then RaiseNativeFailure "SendRequest", DownloadSource
    If Not WinHttpNativeApi.ReceiveResponse(nativeRequest.Value) Then RaiseNativeFailure "ReceiveResponse", DownloadSource
    ReadResponseMetadata nativeRequest, statusCode, reasonPhrase, responseHeaders, protocolUsed, DownloadSource

    If TryParseContentLength(responseHeaders, contentLength) Then
        contentLengthKnown = True
    Else
        contentLengthKnown = False
    End If
    If HasActiveDecompression(Request) Then
        contentLengthKnown = False
        contentLength = -1
    End If
    If Not contentLengthKnown Then contentLength = -1
    Set result = New HttpDownloadResult
    result.Initialize statusCode, reasonPhrase, responseHeaders, protocolUsed, destination, contentLength, contentLengthKnown, HttpTiming.ElapsedMilliseconds(started, HttpTiming.CounterValue())
    If Not result.IsSuccess Then
        Set DownloadCore = result
        CloseNativeHandles nativeRequest, connection, session
        Exit Function
    End If

    temporaryPath = WinHttpDownloadFileSystem.CreateTemporaryPath(destination)
    Set writer = New WinHttpDownloadFileWriter
    writer.OpenNew temporaryPath
    ReDim chunk(0 To ReadChunkBytes - 1)
    NotifyProgress Progress, bytesWritten, ProgressTotal(contentLength, contentLengthKnown)
    Do
        CheckDownloadControls Options, started
        If Not WinHttpNativeApi.QueryDataAvailable(nativeRequest.Value, available) Then RaiseNativeFailure "QueryDataAvailable", DownloadSource
        If available < 0 Then HttpErrors.RaiseTransport HttpErrorIo, DownloadSource, "Native response chunk size exceeded the VBA Long range."
        If available = 0 Then Exit Do
        requested = available
        If requested > ReadChunkBytes Then requested = ReadChunkBytes
        If Not WinHttpNativeApi.ReadData(nativeRequest.Value, chunk, requested, readBytes) Then RaiseNativeFailure "ReadData", DownloadSource
        If readBytes = 0 Then Exit Do
        If readBytes < ReadChunkBytes Then
            ReDim finalChunk(0 To readBytes - 1)
            For copyIndex = 0 To readBytes - 1
                finalChunk(copyIndex) = chunk(copyIndex)
            Next copyIndex
            writer.WriteChunk finalChunk
            Erase finalChunk
        Else
            writer.WriteChunk chunk
        End If
        bytesWritten = AddByteCount(bytesWritten, readBytes)
        CheckDownloadControls Options, started
        NotifyProgress Progress, bytesWritten, ProgressTotal(contentLength, contentLengthKnown)
    Loop

    CheckDownloadControls Options, started
    writer.FlushAndClose
    WinHttpDownloadFileSystem.ReplaceFileAtomically temporaryPath, destination
    temporaryPath = ""
    result.MarkPublished bytesWritten
    Set DownloadCore = result
    CloseNativeHandles nativeRequest, connection, session
    Exit Function

DownloadFailure:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo CleanupFailed
    If Not writer Is Nothing Then cleanupSucceeded = writer.Abort
    If Not cleanupSucceeded Then Err.Clear
    If Len(temporaryPath) > 0 Then WinHttpDownloadFileSystem.DeleteFileIfExists temporaryPath
    CloseNativeHandles nativeRequest, connection, session
    On Error GoTo 0
    GoTo CleanupFinished

CleanupFailed:
    Err.Clear
    On Error GoTo 0

CleanupFinished:

    If errorNumber <> 0 Then
        If HttpErrors.CategoryFromNumber(errorNumber) <> HttpErrorNone Then
            Err.Raise errorNumber, errorSource, errorDescription
        Else
            WinHttpErrorMapping.RaiseMappedWinHttpFailure errorNumber, DownloadSource
        End If
        Err.Raise errorNumber, errorSource, errorDescription
    Else
        Err.Raise HttpErrIo, DownloadSource, "Native streaming download failed without an error number."
    End If
End Function

Private Function UploadFileCore(ByVal Request As HttpRequest, ByVal SourcePath As String, ByVal ContentType As String, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing) As HttpUploadResult
    Dim concreteSource As New WinHttpFileUploadSource
    Dim source As IHttpUploadSource

    concreteSource.Initialize SourcePath, ContentType
    Set source = concreteSource
    Set UploadFileCore = UploadCore(Request, source, Options, Progress, UploadSource)
End Function

Private Function UploadMultipartCore(ByVal Request As HttpRequest, ByVal Form As HttpMultipartForm, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing) As HttpUploadResult
    Dim concreteSource As New WinHttpMultipartUploadSource
    Dim source As IHttpUploadSource

    concreteSource.Initialize Form
    Set source = concreteSource
    Set UploadMultipartCore = UploadCore(Request, source, Options, Progress, UploadSource)
End Function

Private Function UploadCore(ByVal Request As HttpRequest, ByVal Source As IHttpUploadSource, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing, Optional ByVal SourceName As String = "") As HttpUploadResult
    Dim Url As WinHttpNativeUrl
    Dim session As WinHttpSessionHandle
    Dim connection As WinHttpConnectionHandle
    Dim nativeRequest As WinHttpRequestHandle
    Dim writer As WinHttpUploadStreamWriter
    Dim result As HttpUploadResult
    Dim headers As HttpHeaders
    Dim statusCode As Long
    Dim reasonPhrase As String
    Dim responseHeaders As HttpHeaders
    Dim protocolUsed As String
    Dim contentLength As Currency
    Dim started As Currency
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    #If VBA7 Then
    Dim handleValue As LongPtr
    #Else
    Dim handleValue As Long
    #End If

    If Len(SourceName) = 0 Then SourceName = UploadSource
    On Error GoTo UploadFailure
    If Request Is Nothing Then HttpErrors.RaiseValidation SourceName, "Request cannot be Nothing."
    If Source Is Nothing Then HttpErrors.RaiseValidation SourceName, "Upload source cannot be Nothing."
    If Options Is Nothing Then HttpErrors.RaiseValidation SourceName, "Execution options cannot be Nothing."
    If Request.Method <> "POST" Then HttpErrors.RaiseValidation SourceName, "Streaming uploads require the POST method."
    If Request.FollowRedirects Then HttpErrors.RaiseValidation SourceName, "Streaming uploads do not follow redirects."
    If Request.Headers.Contains("Transfer-Encoding") Then HttpErrors.RaiseValidation SourceName, "Streaming uploads manage transfer encoding internally."

    contentLength = Source.ContentLength
    If contentLength < 0 Then HttpErrors.RaiseValidation SourceName, "Upload content length cannot be negative."
    CheckUploadControls Options, 0, SourceName
    started = HttpTiming.CounterValue()
    Set Url = New WinHttpNativeUrl
    Url.Initialize Request.Url
    Set session = New WinHttpSessionHandle
    handleValue = OpenSessionForRequest(Request, SourceName)
    If handleValue = 0 Then RaiseNativeFailure "OpenSession", SourceName
    session.Initialize handleValue

    Set connection = New WinHttpConnectionHandle
    handleValue = WinHttpNativeApi.Connect(session.Value, Url.Host, Url.Port)
    If handleValue = 0 Then RaiseNativeFailure "Connect", SourceName
    connection.Initialize handleValue, session

    Set nativeRequest = New WinHttpRequestHandle
    If Url.IsSecure Then
        handleValue = WinHttpNativeApi.OpenRequest(connection.Value, Request.Method, Url.PathAndQuery, WinHttpNativeApi.WinHttpFlagSecure)
    Else
        handleValue = WinHttpNativeApi.OpenRequest(connection.Value, Request.Method, Url.PathAndQuery, 0)
    End If
    If handleValue = 0 Then RaiseNativeFailure "OpenRequest", SourceName
    nativeRequest.Initialize handleValue, connection

    ApplyRequestOptions nativeRequest, Request, SourceName
    Set headers = Request.Headers.Clone()
    headers.SetValue "Content-Length", WinHttpUploadEncoding.ByteCountText(contentLength)
    headers.SetValue "Content-Type", Source.ContentType
    If Not WinHttpNativeApi.AddRequestHeaders(nativeRequest.Value, BuildRequestHeaders(headers)) Then RaiseNativeFailure "AddRequestHeaders", SourceName
    If Not WinHttpNativeApi.BeginUpload(nativeRequest.Value, contentLength) Then RaiseNativeFailure "SendRequest", SourceName

    Set writer = New WinHttpUploadStreamWriter
    writer.Initialize nativeRequest, Options, Progress, contentLength, SourceName
    writer.SetStarted started
    NotifyUploadProgress Progress, 0, contentLength
    Source.WriteTo writer
    If writer.BytesWritten <> contentLength Then HttpErrors.RaiseTransport HttpErrorIo, SourceName, "Upload source length changed while writing."
    CheckUploadControls Options, started, SourceName
    If Not WinHttpNativeApi.ReceiveResponse(nativeRequest.Value) Then RaiseNativeFailure "ReceiveResponse", SourceName
    ReadResponseMetadata nativeRequest, statusCode, reasonPhrase, responseHeaders, protocolUsed, SourceName

    Set result = New HttpUploadResult
    result.Initialize statusCode, reasonPhrase, responseHeaders, protocolUsed, writer.BytesWritten, contentLength, HttpTiming.ElapsedMilliseconds(started, HttpTiming.CounterValue())
    Set UploadCore = result
    CloseNativeHandles nativeRequest, connection, session
    Exit Function

UploadFailure:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    CloseNativeHandles nativeRequest, connection, session
    If errorNumber <> 0 Then
        If HttpErrors.CategoryFromNumber(errorNumber) <> HttpErrorNone Then
            Err.Raise errorNumber, errorSource, errorDescription
        Else
            WinHttpErrorMapping.RaiseMappedWinHttpFailure errorNumber, SourceName
        End If
        Err.Raise errorNumber, errorSource, errorDescription
    Else
        Err.Raise HttpErrIo, SourceName, "Native streaming upload failed without an error number."
    End If
End Function

Private Sub CloseNativeHandles(ByVal nativeRequest As WinHttpRequestHandle, ByVal connection As WinHttpConnectionHandle, ByVal session As WinHttpSessionHandle)
    If Not nativeRequest Is Nothing Then nativeRequest.ReleaseHandle
    If Not connection Is Nothing Then connection.ReleaseHandle
    If Not session Is Nothing Then session.ReleaseHandle
End Sub

Private Sub RaiseNativeFailure(Optional ByVal Operation As String = "", Optional ByVal Source As String = "")
    If Len(Source) = 0 Then Source = TransportSource
    WinHttpErrorMapping.RaiseMappedWinHttpFailure WinHttpNativeApi.LastErrorCode(), Source, Operation
End Sub

#If VBA7 Then
Private Function OpenSessionForRequest(ByVal Request As HttpRequest, ByVal Source As String) As LongPtr
    Dim Options As HttpProxyOptions
    Dim accessType As Long

    If Request Is Nothing Then HttpErrors.RaiseValidation Source, "Request cannot be Nothing."
    Set Options = Request.ProxyOptions
    If Options Is Nothing Then
        accessType = WinHttpNativeApi.WinHttpAccessTypeDefaultProxy
    Else
        Options.Validate
        Select Case Options.Mode
        Case HttpProxyDefault
            accessType = WinHttpNativeApi.WinHttpAccessTypeDefaultProxy
        Case HttpProxyNoProxy
            accessType = WinHttpNativeApi.WinHttpAccessTypeNoProxy
        Case HttpProxyManual
            accessType = WinHttpNativeApi.WinHttpAccessTypeNamedProxy
        Case Else
            HttpErrors.RaiseValidation Source, "Unsupported proxy mode."
        End Select
    End If

    If Options Is Nothing Then
        OpenSessionForRequest = WinHttpNativeApi.OpenSession(UserAgent, accessType)
    Else
        OpenSessionForRequest = WinHttpNativeApi.OpenSession(UserAgent, accessType, ProxyServerName(Options.ProxyUrl), Options.BypassList)
    End If
End Function
#Else
Private Function OpenSessionForRequest(ByVal Request As HttpRequest, ByVal Source As String) As Long
    Dim Options As HttpProxyOptions
    Dim accessType As Long

    If Request Is Nothing Then HttpErrors.RaiseValidation Source, "Request cannot be Nothing."
    Set Options = Request.ProxyOptions
    If Options Is Nothing Then
        accessType = WinHttpNativeApi.WinHttpAccessTypeDefaultProxy
    Else
        Options.Validate
        Select Case Options.Mode
        Case HttpProxyDefault
            accessType = WinHttpNativeApi.WinHttpAccessTypeDefaultProxy
        Case HttpProxyNoProxy
            accessType = WinHttpNativeApi.WinHttpAccessTypeNoProxy
        Case HttpProxyManual
            accessType = WinHttpNativeApi.WinHttpAccessTypeNamedProxy
        Case Else
            HttpErrors.RaiseValidation Source, "Unsupported proxy mode."
        End Select
    End If

    If Options Is Nothing Then
        OpenSessionForRequest = WinHttpNativeApi.OpenSession(UserAgent, accessType)
    Else
        OpenSessionForRequest = WinHttpNativeApi.OpenSession(UserAgent, accessType, ProxyServerName(Options.ProxyUrl), Options.BypassList)
    End If
End Function
#End If

Private Function ProxyServerName(ByVal ProxyUrl As String) As String
    Dim schemeEnd As Long

    schemeEnd = InStr(1, ProxyUrl, "://", vbBinaryCompare)
    If schemeEnd > 0 Then
        ProxyServerName = Mid$(ProxyUrl, schemeEnd + 3)
    Else
        ProxyServerName = ProxyUrl
    End If
End Function

Private Sub ApplyRequestOptions(ByVal nativeRequest As WinHttpRequestHandle, ByVal Request As HttpRequest, Optional ByVal Source As String = "")
    Dim redirectFlag As Long

    If Not WinHttpNativeApi.SetTimeouts(nativeRequest.Value, Request.Timeouts.ResolveMilliseconds, Request.Timeouts.ConnectMilliseconds, Request.Timeouts.SendMilliseconds, Request.Timeouts.ReceiveMilliseconds) Then RaiseNativeFailure "SetTimeouts", Source
    If Request.FollowRedirects And Not HttpSecurity.HasRedirectSensitiveHeaders(Request.Headers) Then
        ' Keep WinHTTP's secure default: follow redirects except HTTPS -> HTTP.
        redirectFlag = WinHttpNativeApi.WinHttpRedirectPolicyDisallowHttpsToHttp
    Else
        redirectFlag = WinHttpNativeApi.WinHttpRedirectPolicyNever
    End If
    If Not WinHttpNativeApi.SetOptionLong(nativeRequest.Value, WinHttpNativeApi.WinHttpOptionRedirectPolicy, redirectFlag) Then RaiseNativeFailure "SetRedirects", Source
    If Not WinHttpNativeApi.SetOptionLong(nativeRequest.Value, WinHttpNativeApi.WinHttpOptionMaxAutomaticRedirects, Request.MaxRedirects) Then RaiseNativeFailure "SetMaxRedirects", Source
    ApplyProtocolOptions nativeRequest, Request.Url, Request.ProtocolOptions, Source
    ApplyDecompressionOptions nativeRequest, Request.Headers, Request.DecompressionOptions, Source
End Sub

Private Sub ApplyProtocolOptions(ByVal nativeRequest As WinHttpRequestHandle, ByVal Url As String, ByVal Options As HttpProtocolOptions, ByVal Source As String)
    Dim errorCode As Long

    If Options Is Nothing Then Exit Sub
    Options.Validate
    If Not Options.HasOverride Then Exit Sub
    If LCase$(Left$(Url, 8)) <> "https://" Then
        If Options.Mode = HttpProtocolRequired Then
            HttpErrors.RaiseTransport HttpErrorProtocol, Source, "HTTP/2 or HTTP/3 requires an HTTPS request URL."
        End If
        Exit Sub
    End If
    If Not WinHttpNativeApi.SetOptionLong(nativeRequest.Value, WinHttpNativeApi.WinHttpOptionEnableHttpProtocol, Options.EnabledProtocols) Then
        errorCode = WinHttpNativeApi.LastErrorCode()
        If Options.Mode = HttpProtocolAllowFallback And errorCode = WinHttpNativeApi.WinHttpErrorInvalidOption Then Exit Sub
        If Options.Mode = HttpProtocolRequired And errorCode = WinHttpNativeApi.WinHttpErrorInvalidOption Then
            HttpErrors.RaiseTransport HttpErrorProtocol, Source, "Requested HTTP protocol is not supported by this WinHTTP runtime."
        End If
        RaiseNativeFailure "SetHttpProtocol", Source
    End If
    If Options.Mode = HttpProtocolRequired Then
        If Not WinHttpNativeApi.SetOptionLong(nativeRequest.Value, WinHttpNativeApi.WinHttpOptionHttpProtocolRequired, 1) Then
            errorCode = WinHttpNativeApi.LastErrorCode()
            If errorCode = WinHttpNativeApi.WinHttpErrorInvalidOption Then
                HttpErrors.RaiseTransport HttpErrorProtocol, Source, "Required HTTP protocol mode is not supported by this WinHTTP runtime."
            End If
            RaiseNativeFailure "RequireHttpProtocol", Source
        End If
    End If
End Sub

Private Sub ApplyDecompressionOptions(ByVal nativeRequest As WinHttpRequestHandle, ByVal Headers As HttpHeaders, ByVal Options As HttpDecompressionOptions, ByVal Source As String)
    Dim errorCode As Long

    If Len(Source) = 0 Then Source = TransportSource
    If Options Is Nothing Then Exit Sub
    Options.Validate
    If Not Options.HasOverride Then Exit Sub
    If Not Headers Is Nothing Then
        If Headers.Contains("Accept-Encoding") Then
            HttpErrors.RaiseValidation Source, "Accept-Encoding cannot be supplied when response decompression is enabled."
        End If
    End If
    If Not WinHttpNativeApi.SetOptionLong(nativeRequest.Value, WinHttpNativeApi.WinHttpOptionDecompression, Options.EnabledEncodings) Then
        errorCode = WinHttpNativeApi.LastErrorCode()
        If Options.Mode = HttpDecompressionAllowFallback And errorCode = WinHttpNativeApi.WinHttpErrorInvalidOption Then Exit Sub
        If Options.Mode = HttpDecompressionRequired And errorCode = WinHttpNativeApi.WinHttpErrorInvalidOption Then
            HttpErrors.RaiseTransport HttpErrorProtocol, Source, "Response decompression is not supported by this WinHTTP runtime."
        End If
        RaiseNativeFailure "SetDecompression", Source
    End If
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
    Dim statusCode As Long
    Dim reasonPhrase As String
    Dim responseHeaders As HttpHeaders
    Dim Body As HttpBody
    Dim protocolUsed As String
    Dim response As New HttpResponse

    ReadResponseMetadata nativeRequest, statusCode, reasonPhrase, responseHeaders, protocolUsed, TransportSource
    Set Body = ReadBufferedBody(nativeRequest)

    response.Initialize statusCode, reasonPhrase, responseHeaders, Body, elapsedMilliseconds, protocolUsed
    Set BuildResponse = response
End Function

Private Sub ReadResponseMetadata(ByVal nativeRequest As WinHttpRequestHandle, ByRef statusCode As Long, ByRef reasonPhrase As String, ByRef responseHeaders As HttpHeaders, ByRef protocolUsed As String, Optional ByVal Source As String = "")
    Dim statusText As String
    Dim rawHeaders As String

    If Len(Source) = 0 Then Source = TransportSource
    If Not WinHttpNativeApi.QueryHeaderString(nativeRequest.Value, WinHttpNativeApi.WinHttpQueryStatusCode, statusText) Then RaiseNativeFailure "QueryStatusCode", Source
    On Error GoTo InvalidStatus
    statusCode = CLng(Trim$(statusText))
    On Error GoTo 0
    If statusCode < 100 Or statusCode > 599 Then GoTo InvalidStatus
    If Not WinHttpNativeApi.QueryHeaderString(nativeRequest.Value, WinHttpNativeApi.WinHttpQueryStatusText, reasonPhrase) Then RaiseNativeFailure "QueryStatusText", Source
    If Not WinHttpNativeApi.QueryHeaderString(nativeRequest.Value, WinHttpNativeApi.WinHttpQueryRawHeadersCrlf, rawHeaders) Then RaiseNativeFailure "QueryRawHeaders", Source
    Set responseHeaders = ParseResponseHeaders(rawHeaders)
    protocolUsed = QueryProtocol(nativeRequest, Source)
    Exit Sub

InvalidStatus:
    On Error GoTo 0
    Err.Raise HttpErrProtocol, Source, "WinHTTP returned an invalid response status."
End Sub

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

Private Function QueryProtocol(ByVal nativeRequest As WinHttpRequestHandle, Optional ByVal Source As String = "") As String
    Dim protocolFlags As Long

    If Len(Source) = 0 Then Source = TransportSource
    If Not WinHttpNativeApi.QueryProtocolUsed(nativeRequest.Value, protocolFlags) Then
        If WinHttpNativeApi.LastErrorCode() <> WinHttpNativeApi.WinHttpErrorInvalidOption Then RaiseNativeFailure "QueryProtocolUsed", Source
    End If
    If (protocolFlags And WinHttpNativeApi.WinHttpProtocolFlagHttp3) <> 0 Then
        QueryProtocol = "HTTP/3"
    ElseIf (protocolFlags And WinHttpNativeApi.WinHttpProtocolFlagHttp2) <> 0 Then
        QueryProtocol = "HTTP/2"
    Else
        QueryProtocol = "HTTP/1.1"
    End If
End Function

Private Sub CheckDownloadControls(ByVal Options As HttpExecutionOptions, ByVal started As Currency)
    If Options Is Nothing Then Exit Sub
    If Not Options.CancellationToken Is Nothing Then
        If Options.CancellationToken.IsCancellationRequested Then Err.Raise HttpErrCancelled, DownloadSource, "HTTP download was cancelled."
    End If
    If Options.TotalDeadlineMilliseconds > 0 And started <> 0 Then
        If HttpTiming.ElapsedMilliseconds(started, HttpTiming.CounterValue()) >= Options.TotalDeadlineMilliseconds Then
            Err.Raise HttpErrTimeout, DownloadSource, "HTTP download total deadline expired."
        End If
    End If
End Sub

Private Sub CheckUploadControls(ByVal Options As HttpExecutionOptions, ByVal started As Currency, ByVal Source As String)
    If Options Is Nothing Then Exit Sub
    If Not Options.CancellationToken Is Nothing Then
        If Options.CancellationToken.IsCancellationRequested Then Err.Raise HttpErrCancelled, Source, "HTTP upload was cancelled."
    End If
    If Options.TotalDeadlineMilliseconds > 0 And started <> 0 Then
        If HttpTiming.ElapsedMilliseconds(started, HttpTiming.CounterValue()) >= Options.TotalDeadlineMilliseconds Then Err.Raise HttpErrTimeout, Source, "HTTP upload total deadline expired."
    End If
End Sub

Private Sub NotifyProgress(ByVal Progress As IHttpProgressSink, ByVal bytesWritten As Currency, ByVal totalBytes As Currency)
    If Progress Is Nothing Then Exit Sub
    Progress.OnProgress bytesWritten, totalBytes
End Sub

Private Sub NotifyUploadProgress(ByVal Progress As IHttpProgressSink, ByVal bytesWritten As Currency, ByVal totalBytes As Currency)
    If Not Progress Is Nothing Then Progress.OnProgress bytesWritten, totalBytes
End Sub

Private Function ProgressTotal(ByVal contentLength As Currency, ByVal contentLengthKnown As Boolean) As Currency
    If contentLengthKnown Then
        ProgressTotal = contentLength
    Else
        ProgressTotal = -1
    End If
End Function

Private Function HasActiveDecompression(ByVal Request As HttpRequest) As Boolean
    If Request Is Nothing Then Exit Function
    If Request.DecompressionOptions Is Nothing Then Exit Function
    HasActiveDecompression = Request.DecompressionOptions.HasOverride
End Function

Private Function AddByteCount(ByVal currentValue As Currency, ByVal increment As Long) As Currency
    On Error GoTo Overflow
    AddByteCount = currentValue + CCur(increment)
    Exit Function

Overflow:
    On Error GoTo 0
    Err.Raise HttpErrIo, DownloadSource, "Streaming download byte count exceeded the supported Currency range."
End Function

Private Function TryParseContentLength(ByVal Headers As HttpHeaders, ByRef contentLength As Currency) As Boolean
    Dim rawValue As String

    contentLength = -1
    If Headers Is Nothing Then Exit Function
    If Not Headers.Contains("Content-Length") Then Exit Function
    rawValue = Trim$(Headers.GetValue("Content-Length"))
    If Len(rawValue) = 0 Or Not IsDigits(rawValue) Then Exit Function
    On Error GoTo InvalidLength
    contentLength = CCur(rawValue)
    If contentLength < 0 Then GoTo InvalidLength
    TryParseContentLength = True
    Exit Function

InvalidLength:
    contentLength = -1
    TryParseContentLength = False
End Function

Private Function IsDigits(ByVal value As String) As Boolean
    Dim index As Long
    Dim code As Long

    If Len(value) = 0 Then Exit Function
    For index = 1 To Len(value)
        code = AscW(Mid$(value, index, 1))
        If code < 48 Or code > 57 Then Exit Function
    Next index
    IsDigits = True
End Function
