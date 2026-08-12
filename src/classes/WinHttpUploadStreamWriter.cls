VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpUploadStreamWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mRequest As WinHttpRequestHandle
Private mOptions As HttpExecutionOptions
Private mProgress As IHttpProgressSink
Private mTotalBytes As Currency
Private mBytesWritten As Currency
Private mSource As String
Private mStarted As Currency

Friend Sub Initialize(ByVal Request As WinHttpRequestHandle, ByVal Options As HttpExecutionOptions, ByVal Progress As IHttpProgressSink, ByVal TotalBytes As Currency, ByVal Source As String)
    If Request Is Nothing Then HttpErrors.RaiseValidation "WinHttpUploadStreamWriter.Initialize", "Request cannot be Nothing."
    If Options Is Nothing Then HttpErrors.RaiseValidation "WinHttpUploadStreamWriter.Initialize", "Options cannot be Nothing."
    If TotalBytes < 0 Then HttpErrors.RaiseValidation "WinHttpUploadStreamWriter.Initialize", "Total bytes cannot be negative."

    Set mRequest = Request
    Set mOptions = Options
    Set mProgress = Progress
    mTotalBytes = TotalBytes
    mBytesWritten = 0
    mSource = Source
End Sub

Friend Sub WriteBytes(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal ByteCount As Long)
    Dim remaining As Long
    Dim cursor As Long
    Dim written As Long

    If mRequest Is Nothing Then HttpErrors.RaiseValidation "WinHttpUploadStreamWriter.WriteBytes", "Writer is not initialized."
    If ByteCount < 0 Or Offset < 0 Then HttpErrors.RaiseValidation "WinHttpUploadStreamWriter.WriteBytes", "Buffer range cannot be negative."
    If ByteCount = 0 Then Exit Sub

    remaining = ByteCount
    cursor = Offset
    Do While remaining > 0
        CheckControls
        written = 0
        If Not WinHttpNativeApi.WriteData(mRequest.Value, Buffer, cursor, remaining, written) Then
            WinHttpErrorMapping.RaiseMappedWinHttpFailure WinHttpNativeApi.LastErrorCode(), mSource, "WriteData"
        End If
        If written <= 0 Or written > remaining Then HttpErrors.RaiseTransport HttpErrorIo, mSource, "WinHTTP wrote an invalid byte count."
        mBytesWritten = AddByteCount(mBytesWritten, written)
        remaining = remaining - written
        cursor = cursor + written
        CheckControls
        NotifyProgress
    Loop
End Sub

Friend Property Get BytesWritten() As Currency
    BytesWritten = mBytesWritten
End Property

Private Sub CheckControls()
    If Not mOptions.CancellationToken Is Nothing Then
        If mOptions.CancellationToken.IsCancellationRequested Then Err.Raise HttpErrCancelled, mSource, "HTTP upload was cancelled."
    End If
    If mOptions.TotalDeadlineMilliseconds > 0 Then
        If HttpTiming.ElapsedMilliseconds(mStarted, HttpTiming.CounterValue()) >= mOptions.TotalDeadlineMilliseconds Then Err.Raise HttpErrTimeout, mSource, "HTTP upload total deadline expired."
    End If
End Sub

Private Sub NotifyProgress()
    If Not mProgress Is Nothing Then mProgress.OnProgress mBytesWritten, mTotalBytes
End Sub

Private Function AddByteCount(ByVal currentValue As Currency, ByVal increment As Long) As Currency
    On Error GoTo Overflow
    AddByteCount = currentValue + CCur(increment)
    Exit Function

Overflow:
    On Error GoTo 0
    HttpErrors.RaiseTransport HttpErrorIo, mSource, "Upload byte count overflowed the supported range."
End Function

Friend Sub SetStarted(ByVal Started As Currency)
    mStarted = Started
End Sub
