VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpDiagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DefaultMaxEvents As Long = 100
Private Const MaximumMaxEvents As Long = 10000

Private mEnabled As Boolean
Private mMaxEvents As Long
Private mNextSequence As Long
Private mEvents As New Collection

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal value As Boolean)
    mEnabled = value
End Property

Public Property Get MaxEvents() As Long
    MaxEvents = mMaxEvents
End Property

Public Property Let MaxEvents(ByVal value As Long)
    If value < 1 Or value > MaximumMaxEvents Then
        HttpErrors.RaiseValidation "HttpDiagnostics.MaxEvents", "MaxEvents must be between 1 and 10000."
    End If
    mMaxEvents = value
    TrimEvents
End Property

Public Property Get Count() As Long
    Count = mEvents.Count
End Property

Public Function ItemAt(ByVal index As Long) As HttpDiagnosticEvent
    If index < 1 Or index > mEvents.Count Then HttpErrors.RaiseValidation "HttpDiagnostics.ItemAt", "Diagnostic event index is out of range."
    Set ItemAt = mEvents(index)
End Function

Public Property Get LastEvent() As HttpDiagnosticEvent
    If mEvents.Count > 0 Then Set LastEvent = mEvents(mEvents.Count)
End Property

Public Sub Clear()
    Set mEvents = New Collection
    mNextSequence = 0
End Sub

Public Function ToJson() As String
    Dim index As Long
    Dim result As String

    result = "{" & Chr$(34) & "schema_version" & Chr$(34) & ":1," & _
        Chr$(34) & "enabled" & Chr$(34) & ":" & HttpJson.BooleanValue(mEnabled) & "," & _
        Chr$(34) & "events" & Chr$(34) & ":["
    For index = 1 To mEvents.Count
        If index > 1 Then result = result & ","
        result = result & mEvents(index).ToJson
    Next index
    ToJson = result & "]}"
End Function

Friend Sub RecordResponse(ByVal operation As String, ByVal Request As HttpRequest, ByVal Response As HttpResponse)
    If Response Is Nothing Then Exit Sub
    AppendEvent operation, Request, Response.StatusCode, Response.Headers, Response.ElapsedMilliseconds, Response.ProtocolUsed, 0
End Sub

Friend Sub RecordDownload(ByVal operation As String, ByVal Request As HttpRequest, ByVal Result As HttpDownloadResult)
    If Result Is Nothing Then Exit Sub
    AppendEvent operation, Request, Result.StatusCode, Result.Headers, Result.ElapsedMilliseconds, Result.ProtocolUsed, 0
End Sub

Friend Sub RecordUpload(ByVal operation As String, ByVal Request As HttpRequest, ByVal Result As HttpUploadResult)
    If Result Is Nothing Then Exit Sub
    AppendEvent operation, Request, Result.StatusCode, Result.Headers, Result.ElapsedMilliseconds, Result.ProtocolUsed, 0
End Sub

Friend Sub RecordFailure(ByVal operation As String, ByVal Request As HttpRequest, ByVal errorNumber As Long)
    AppendEvent operation, Request, 0, Nothing, 0, "", errorNumber
End Sub

Private Sub AppendEvent(ByVal operation As String, ByVal Request As HttpRequest, ByVal statusCode As Long, ByVal ResponseHeaders As HttpHeaders, ByVal elapsedMilliseconds As Double, ByVal protocolUsed As String, ByVal errorNumber As Long)
    Dim item As New HttpDiagnosticEvent

    If Not mEnabled Then Exit Sub
    mNextSequence = mNextSequence + 1
    item.Initialize mNextSequence, operation, Request, statusCode, ResponseHeaders, elapsedMilliseconds, protocolUsed, errorNumber
    mEvents.Add item
    TrimEvents
End Sub

Private Sub TrimEvents()
    If mEvents Is Nothing Then Exit Sub
    Do While mEvents.Count > mMaxEvents
        mEvents.Remove 1
    Loop
End Sub

Private Sub Class_Initialize()
    Set mEvents = New Collection
    mMaxEvents = DefaultMaxEvents
End Sub
