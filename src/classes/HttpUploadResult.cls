VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpUploadResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mStatusCode As Long
Private mReasonPhrase As String
Private mHeaders As HttpHeaders
Private mProtocolUsed As String
Private mBytesWritten As Currency
Private mContentLength As Currency
Private mElapsedMilliseconds As Double
Private mAuthenticationChallenged As Boolean

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

Public Property Get BytesWritten() As Currency
    BytesWritten = mBytesWritten
End Property

Public Property Get ContentLength() As Currency
    ContentLength = mContentLength
End Property

Public Property Get ElapsedMilliseconds() As Double
    ElapsedMilliseconds = mElapsedMilliseconds
End Property

Public Property Get AuthenticationChallenged() As Boolean
    AuthenticationChallenged = mAuthenticationChallenged
End Property

Public Property Get IsSuccess() As Boolean
    IsSuccess = (mStatusCode >= 200 And mStatusCode <= 299)
End Property

Public Sub RaiseForStatus()
    If Not IsSuccess Then HttpErrors.RaiseStatus "HttpUploadResult.RaiseForStatus", mStatusCode
End Sub

Friend Sub Initialize(ByVal statusCode As Long, ByVal reasonPhrase As String, ByVal Headers As HttpHeaders, ByVal protocolUsed As String, ByVal bytesWritten As Currency, ByVal contentLength As Currency, ByVal elapsedMilliseconds As Double)
    If statusCode < 100 Or statusCode > 599 Then
        HttpErrors.RaiseValidation "HttpUploadResult.Initialize", "HTTP status code must be between 100 and 599."
    End If
    If Headers Is Nothing Then HttpErrors.RaiseValidation "HttpUploadResult.Initialize", "Headers cannot be Nothing."
    If bytesWritten < 0 Or contentLength < 0 Then HttpErrors.RaiseValidation "HttpUploadResult.Initialize", "Upload byte counts cannot be negative."
    If bytesWritten > contentLength Then HttpErrors.RaiseValidation "HttpUploadResult.Initialize", "Bytes written cannot exceed content length."
    If elapsedMilliseconds < 0 Then HttpErrors.RaiseValidation "HttpUploadResult.Initialize", "Elapsed milliseconds cannot be negative."

    mStatusCode = statusCode
    mReasonPhrase = reasonPhrase
    Set mHeaders = Headers.Clone()
    mProtocolUsed = protocolUsed
    mBytesWritten = bytesWritten
    mContentLength = contentLength
    mElapsedMilliseconds = elapsedMilliseconds
    mAuthenticationChallenged = (statusCode = 401 Or statusCode = 407)
End Sub
