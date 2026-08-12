VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpDownloadResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mStatusCode As Long
Private mReasonPhrase As String
Private mHeaders As HttpHeaders
Private mProtocolUsed As String
Private mDestinationPath As String
Private mPublished As Boolean
Private mBytesWritten As Currency
Private mContentLength As Currency
Private mContentLengthKnown As Boolean
Private mElapsedMilliseconds As Double

Public Property Get StatusCode() As Long
    StatusCode = mStatusCode
End Property

Public Property Get ReasonPhrase() As String
    ReasonPhrase = mReasonPhrase
End Property

Public Property Get Headers() As HttpHeaders
    Set Headers = mHeaders
End Property

Public Property Get ProtocolUsed() As String
    ProtocolUsed = mProtocolUsed
End Property

Public Property Get DestinationPath() As String
    DestinationPath = mDestinationPath
End Property

Public Property Get Published() As Boolean
    Published = mPublished
End Property

Public Property Get BytesWritten() As Currency
    BytesWritten = mBytesWritten
End Property

Public Property Get ContentLength() As Currency
    ContentLength = mContentLength
End Property

Public Property Get ContentLengthKnown() As Boolean
    ContentLengthKnown = mContentLengthKnown
End Property

Public Property Get ElapsedMilliseconds() As Double
    ElapsedMilliseconds = mElapsedMilliseconds
End Property

Public Property Get IsSuccess() As Boolean
    IsSuccess = (mStatusCode >= 200 And mStatusCode <= 299)
End Property

Public Sub RaiseForStatus()
    If Not IsSuccess Then HttpErrors.RaiseStatus "HttpDownloadResult.RaiseForStatus", mStatusCode
End Sub

Friend Sub Initialize(ByVal statusCode As Long, ByVal reasonPhrase As String, ByVal Headers As HttpHeaders, ByVal protocolUsed As String, ByVal destinationPath As String, ByVal contentLength As Currency, ByVal contentLengthKnown As Boolean, ByVal elapsedMilliseconds As Double)
    If statusCode < 100 Or statusCode > 599 Then
        HttpErrors.RaiseValidation "HttpDownloadResult.Initialize", "HTTP status code must be between 100 and 599."
    End If
    If Headers Is Nothing Then HttpErrors.RaiseValidation "HttpDownloadResult.Initialize", "Headers cannot be Nothing."
    If contentLengthKnown And contentLength < 0 Then HttpErrors.RaiseValidation "HttpDownloadResult.Initialize", "Known content length cannot be negative."
    If Not contentLengthKnown Then contentLength = -1

    mStatusCode = statusCode
    mReasonPhrase = reasonPhrase
    Set mHeaders = Headers.Clone()
    mProtocolUsed = protocolUsed
    mDestinationPath = destinationPath
    mPublished = False
    mBytesWritten = 0
    mContentLength = contentLength
    mContentLengthKnown = contentLengthKnown
    mElapsedMilliseconds = elapsedMilliseconds
End Sub

Friend Sub MarkPublished(ByVal bytesWritten As Currency)
    If bytesWritten < 0 Then HttpErrors.RaiseValidation "HttpDownloadResult.MarkPublished", "Bytes written cannot be negative."
    If Not IsSuccess Then HttpErrors.RaiseValidation "HttpDownloadResult.MarkPublished", "Only a successful HTTP result can be published."
    mBytesWritten = bytesWritten
    mPublished = True
End Sub
