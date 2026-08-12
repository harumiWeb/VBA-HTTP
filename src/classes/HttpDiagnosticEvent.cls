VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpDiagnosticEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mSequence As Long
Private mOperation As String
Private mMethod As String
Private mTarget As String
Private mStatusCode As Long
Private mProtocolUsed As String
Private mElapsedMilliseconds As Double
Private mErrorNumber As Long
Private mErrorCategory As HttpErrorCategory
Private mRequestHeaders As HttpHeaders
Private mResponseHeaders As HttpHeaders

Public Property Get Sequence() As Long
    Sequence = mSequence
End Property

Public Property Get Operation() As String
    Operation = mOperation
End Property

Public Property Get Method() As String
    Method = mMethod
End Property

Public Property Get Target() As String
    Target = mTarget
End Property

Public Property Get StatusCode() As Long
    StatusCode = mStatusCode
End Property

Public Property Get ProtocolUsed() As String
    ProtocolUsed = mProtocolUsed
End Property

Public Property Get ElapsedMilliseconds() As Double
    ElapsedMilliseconds = mElapsedMilliseconds
End Property

Public Property Get ErrorNumber() As Long
    ErrorNumber = mErrorNumber
End Property

Public Property Get ErrorCategory() As HttpErrorCategory
    ErrorCategory = mErrorCategory
End Property

Public Property Get RequestHeaders() As HttpHeaders
    If mRequestHeaders Is Nothing Then
        Set RequestHeaders = New HttpHeaders
    Else
        Set RequestHeaders = mRequestHeaders.Clone()
    End If
End Property

Public Property Get ResponseHeaders() As HttpHeaders
    If mResponseHeaders Is Nothing Then
        Set ResponseHeaders = New HttpHeaders
    Else
        Set ResponseHeaders = mResponseHeaders.Clone()
    End If
End Property

Public Function ToJson() As String
    ToJson = "{" & _
        Chr$(34) & "schema_version" & Chr$(34) & ":1," & _
        Chr$(34) & "sequence" & Chr$(34) & ":" & CStr(mSequence) & "," & _
        Chr$(34) & "operation" & Chr$(34) & ":" & HttpJson.Quote(mOperation) & "," & _
        Chr$(34) & "method" & Chr$(34) & ":" & HttpJson.Quote(mMethod) & "," & _
        Chr$(34) & "target" & Chr$(34) & ":" & HttpJson.Quote(mTarget) & "," & _
        Chr$(34) & "status_code" & Chr$(34) & ":" & CStr(mStatusCode) & "," & _
        Chr$(34) & "protocol" & Chr$(34) & ":" & HttpJson.Quote(mProtocolUsed) & "," & _
        Chr$(34) & "elapsed_ms" & Chr$(34) & ":" & HttpJson.Number(mElapsedMilliseconds) & "," & _
        Chr$(34) & "error_number" & Chr$(34) & ":" & CStr(mErrorNumber) & "," & _
        Chr$(34) & "error_category" & Chr$(34) & ":" & HttpJson.Quote(HttpErrors.CategoryName(mErrorCategory)) & "," & _
        Chr$(34) & "request_headers" & Chr$(34) & ":" & HttpJson.HeadersArray(mRequestHeaders) & "," & _
        Chr$(34) & "response_headers" & Chr$(34) & ":" & HttpJson.HeadersArray(mResponseHeaders) & "}"
End Function

Friend Sub Initialize(ByVal eventSequence As Long, ByVal eventOperation As String, ByVal Request As HttpRequest, ByVal statusCode As Long, ByVal ResponseHeaders As HttpHeaders, ByVal elapsedMilliseconds As Double, ByVal protocolUsed As String, ByVal errorNumber As Long)
    mSequence = eventSequence
    mOperation = eventOperation
    mMethod = ""
    mTarget = ""
    If Not Request Is Nothing Then
        mMethod = Request.Method
        mTarget = SafeTarget(Request.Url)
        Set mRequestHeaders = HttpSecurity.RedactedHeaders(Request.Headers)
    Else
        Set mRequestHeaders = New HttpHeaders
    End If
    If ResponseHeaders Is Nothing Then
        Set mResponseHeaders = New HttpHeaders
    Else
        Set mResponseHeaders = HttpSecurity.RedactedHeaders(ResponseHeaders)
    End If
    mStatusCode = statusCode
    mProtocolUsed = protocolUsed
    If elapsedMilliseconds < 0 Then elapsedMilliseconds = 0
    mElapsedMilliseconds = elapsedMilliseconds
    mErrorNumber = errorNumber
    mErrorCategory = HttpErrors.CategoryFromNumber(errorNumber)
End Sub

Private Function SafeTarget(ByVal value As String) As String
    Dim schemeEnd As Long
    Dim authorityStart As Long
    Dim authorityEnd As Long
    Dim pathStart As Long
    Dim authority As String
    Dim path As String
    Dim userInfoEnd As Long
    Dim index As Long

    schemeEnd = InStr(1, value, "://", vbBinaryCompare)
    If schemeEnd <= 1 Then
        SafeTarget = "[invalid]"
        Exit Function
    End If
    authorityStart = schemeEnd + 3
    authorityEnd = Len(value) + 1
    For index = authorityStart To Len(value)
        If InStr(1, "/?#", Mid$(value, index, 1), vbBinaryCompare) > 0 Then
            authorityEnd = index
            Exit For
        End If
    Next index
    authority = Mid$(value, authorityStart, authorityEnd - authorityStart)
    userInfoEnd = InStrRev(authority, "@")
    If userInfoEnd > 0 Then authority = Mid$(authority, userInfoEnd + 1)
    If Len(authority) = 0 Then
        SafeTarget = "[invalid]"
        Exit Function
    End If
    If authorityEnd <= Len(value) Then
        pathStart = Len(value) + 1
        For index = authorityEnd To Len(value)
            If InStr(1, "?#", Mid$(value, index, 1), vbBinaryCompare) > 0 Then
                pathStart = index
                Exit For
            End If
        Next index
        path = Mid$(value, authorityEnd, pathStart - authorityEnd)
    End If
    If Len(path) = 0 Then path = "/"
    SafeTarget = Left$(value, schemeEnd - 1) & "://" & authority & path
End Function
