VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mBaseUrl As String
Private mDefaultHeaders As New HttpHeaders
Private mTransport As IHttpTransport
Private mExecuting As Boolean

Public Property Get BaseUrl() As String
    BaseUrl = mBaseUrl
End Property

Public Property Let BaseUrl(ByVal value As String)
    value = Trim$(value)
    If Len(value) > 0 Then ValidateAbsoluteUrl value, "HttpClient.BaseUrl"
    mBaseUrl = value
End Property

Public Property Get DefaultHeaders() As HttpHeaders
    Set DefaultHeaders = mDefaultHeaders
End Property

Public Property Get Transport() As IHttpTransport
    Set Transport = mTransport
End Property

Public Property Set Transport(ByVal value As IHttpTransport)
    If value Is Nothing Then
        HttpErrors.RaiseValidation "HttpClient.Transport", "Transport cannot be Nothing."
    End If
    Set mTransport = value
End Property

Public Function Execute(ByVal Request As HttpRequest) As HttpResponse
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    BeginExecution "HttpClient.Execute"
    On Error GoTo ExecutionFailed
    Set Execute = ExecuteCore(Request)
    mExecuting = False
    Exit Function

ExecutionFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    mExecuting = False
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Public Function ExecuteMany(ByVal Requests As Collection, Optional ByVal Options As HttpBatchOptions = Nothing) As HttpBatchResult
    Dim batchTransport As IHttpBatchTransport
    Dim prepared As Collection
    Dim snapshotOptions As HttpBatchOptions
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    BeginExecution "HttpClient.ExecuteMany"
    On Error GoTo ExecutionFailed
    If Requests Is Nothing Then HttpErrors.RaiseValidation "HttpClient.ExecuteMany", "Requests cannot be Nothing."
    If Not TypeOf mTransport Is IHttpBatchTransport Then HttpErrors.RaiseValidation "HttpClient.ExecuteMany", "Configured transport does not support batch execution."

    Set prepared = PrepareRequests(Requests)
    If Options Is Nothing Then
        Set snapshotOptions = New HttpBatchOptions
    Else
        Set snapshotOptions = Options.Clone()
    End If
    Set batchTransport = mTransport
    Set ExecuteMany = batchTransport.ExecuteMany(prepared, snapshotOptions)
    mExecuting = False
    Exit Function

ExecutionFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    mExecuting = False
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Public Function GetMany(ByVal Urls As Collection, Optional ByVal Options As HttpBatchOptions = Nothing) As HttpBatchResult
    Dim Requests As New Collection
    Dim index As Long
    Dim value As Variant
    Dim Request As HttpRequest

    If Urls Is Nothing Then
        HttpErrors.RaiseValidation "HttpClient.GetMany", "URLs cannot be Nothing."
        Exit Function
    End If
    For index = 1 To Urls.Count
        If IsObject(Urls(index)) Then HttpErrors.RaiseValidation "HttpClient.GetMany", "Every URL must be a non-empty String."
        value = Urls(index)
        If VarType(value) <> vbString Or Len(CStr(value)) = 0 Then
            HttpErrors.RaiseValidation "HttpClient.GetMany", "Every URL must be a non-empty String."
        End If
        Set Request = New HttpRequest
        Request.Method = "GET"
        Request.Url = CStr(value)
        Requests.Add Request
    Next index
    Set GetMany = ExecuteMany(Requests, Options)
End Function

Private Function ExecuteCore(ByVal Request As HttpRequest) As HttpResponse
    Dim snapshot As HttpRequest
    Dim response As HttpResponse

    If Request Is Nothing Then
        HttpErrors.RaiseValidation "HttpClient.Execute", "Request cannot be Nothing."
    End If
    If mTransport Is Nothing Then
        HttpErrors.RaiseValidation "HttpClient.Execute", "No HTTP transport is configured."
    End If

    Set snapshot = PrepareRequest(Request)
    Set response = mTransport.Execute(snapshot)
    If response Is Nothing Then
        Err.Raise HttpErrProtocol, "HttpClient.Execute", "Transport returned Nothing instead of an HttpResponse."
    End If
    Set ExecuteCore = response
End Function

Public Function GetResponse(ByVal Url As String, Optional ByVal Query As HttpParams = Nothing) As HttpResponse
    Dim Request As New HttpRequest

    Request.Method = "GET"
    Request.Url = Url
    If Not Query Is Nothing Then Set Request.Query = Query
    Set GetResponse = Execute(Request)
End Function

Public Function PostResponse(ByVal Url As String, Optional ByVal Body As HttpBody = Nothing) As HttpResponse
    Set PostResponse = ExecuteWithBody("POST", Url, Body)
End Function

Public Function PutResponse(ByVal Url As String, Optional ByVal Body As HttpBody = Nothing) As HttpResponse
    Set PutResponse = ExecuteWithBody("PUT", Url, Body)
End Function

Public Function PatchResponse(ByVal Url As String, Optional ByVal Body As HttpBody = Nothing) As HttpResponse
    Set PatchResponse = ExecuteWithBody("PATCH", Url, Body)
End Function

Public Function DeleteResponse(ByVal Url As String) As HttpResponse
    Dim Request As New HttpRequest

    Request.Method = "DELETE"
    Request.Url = Url
    Set DeleteResponse = Execute(Request)
End Function

Private Function ExecuteWithBody(ByVal Method As String, ByVal Url As String, ByVal Body As HttpBody) As HttpResponse
    Dim Request As New HttpRequest

    Request.Method = Method
    Request.Url = Url
    If Not Body Is Nothing Then Set Request.Body = Body
    Set ExecuteWithBody = Execute(Request)
End Function

Private Function PrepareRequest(ByVal Request As HttpRequest) As HttpRequest
    Dim snapshot As HttpRequest

    If Request Is Nothing Then HttpErrors.RaiseValidation "HttpClient.Execute", "Request cannot be Nothing."
    Set snapshot = Request.Clone()
    snapshot.Url = ResolveUrl(snapshot.Url)
    ApplyQuery snapshot
    ApplyDefaultHeaders snapshot
    Set PrepareRequest = snapshot
End Function

Private Function PrepareRequests(ByVal Requests As Collection) As Collection
    Dim prepared As New Collection
    Dim candidate As Object
    Dim Request As HttpRequest
    Dim index As Long

    If Requests Is Nothing Then
        HttpErrors.RaiseValidation "HttpClient.ExecuteMany", "Requests cannot be Nothing."
        Exit Function
    End If
    For index = 1 To Requests.Count
        If Not IsObject(Requests(index)) Then HttpErrors.RaiseValidation "HttpClient.ExecuteMany", "Every item must be an HttpRequest."
        Set candidate = Requests(index)
        If Not TypeOf candidate Is HttpRequest Then HttpErrors.RaiseValidation "HttpClient.ExecuteMany", "Every item must be an HttpRequest."
        Set Request = candidate
        prepared.Add PrepareRequest(Request)
    Next index
    Set PrepareRequests = prepared
End Function

Private Sub BeginExecution(ByVal Source As String)
    If mExecuting Then HttpErrors.RaiseValidation Source, "The same HttpClient cannot execute reentrantly."
    mExecuting = True
End Sub

Private Function ResolveUrl(ByVal value As String) As String
    If IsAbsoluteHttpUrl(value) Then
        ValidateAbsoluteUrl value, "HttpClient.Execute"
        ResolveUrl = value
        Exit Function
    End If
    If InStr(1, value, "://", vbBinaryCompare) > 0 Then
        HttpErrors.RaiseInvalidUrl "HttpClient.Execute", "Request URL must use HTTP or HTTPS."
    End If
    If Len(mBaseUrl) = 0 Then
        HttpErrors.RaiseInvalidUrl "HttpClient.Execute", "Relative request URL requires BaseUrl."
    End If
    ResolveUrl = RTrimSlash(mBaseUrl) & "/" & LTrimSlash(value)
End Function

Private Sub ValidateAbsoluteUrl(ByVal value As String, ByVal Source As String)
    Dim authorityStart As Long
    Dim authorityEnd As Long
    Dim authority As String
    Dim index As Long
    Dim characterCode As Long

    If Not IsAbsoluteHttpUrl(value) Then
        HttpErrors.RaiseInvalidUrl Source, "URL must be an absolute HTTP or HTTPS URL."
    End If
    For index = 1 To Len(value)
        characterCode = AscW(Mid$(value, index, 1))
        If characterCode <= 32 Or characterCode = 127 Then
            HttpErrors.RaiseInvalidUrl Source, "URL cannot contain whitespace or control characters."
        End If
    Next index

    authorityStart = InStr(1, value, "://", vbBinaryCompare) + 3
    authorityEnd = Len(value) + 1
    For index = authorityStart To Len(value)
        If InStr(1, "/?#", Mid$(value, index, 1), vbBinaryCompare) > 0 Then
            authorityEnd = index
            Exit For
        End If
    Next index
    authority = Mid$(value, authorityStart, authorityEnd - authorityStart)
    If Len(authority) = 0 Then
        HttpErrors.RaiseInvalidUrl Source, "URL must contain a host."
    End If
End Sub

Private Sub ApplyQuery(ByVal Request As HttpRequest)
    Dim queryText As String
    Dim separator As String

    queryText = Request.Query.ToQueryString()
    If Len(queryText) = 0 Then Exit Sub
    If InStr(1, Request.Url, "?", vbBinaryCompare) > 0 Then
        separator = "&"
    Else
        separator = "?"
    End If
    Request.Url = Request.Url & separator & queryText
End Sub

Private Sub ApplyDefaultHeaders(ByVal Request As HttpRequest)
    Dim merged As HttpHeaders
    Dim requestNames As New HttpHeaders
    Dim index As Long
    Dim Name As String

    Set merged = mDefaultHeaders.Clone()
    For index = 1 To Request.Headers.Count
        Name = Request.Headers.NameAt(index)
        If Not requestNames.Contains(Name) Then
            merged.Remove Name
            requestNames.Add Name, ""
        End If
        merged.Add Name, Request.Headers.ValueAt(index)
    Next index
    Set Request.Headers = merged
End Sub

Private Function IsAbsoluteHttpUrl(ByVal value As String) As Boolean
    IsAbsoluteHttpUrl = (LCase$(Left$(value, 7)) = "http://" Or LCase$(Left$(value, 8)) = "https://")
End Function

Private Function RTrimSlash(ByVal value As String) As String
    Do While Right$(value, 1) = "/"
        value = Left$(value, Len(value) - 1)
    Loop
    RTrimSlash = value
End Function

Private Function LTrimSlash(ByVal value As String) As String
    Do While Left$(value, 1) = "/"
        value = Mid$(value, 2)
    Loop
    LTrimSlash = value
End Function

Private Sub Class_Initialize()
    Set mTransport = New WinHttpComTransport
End Sub
