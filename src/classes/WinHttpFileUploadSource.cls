VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpFileUploadSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IHttpUploadSource

Private mPath As String
Private mSize As Currency
Private mContentType As String

Friend Sub Initialize(ByVal SourcePath As String, ByVal ContentType As String)
    If Len(ContentType) = 0 Then HttpErrors.RaiseValidation "WinHttpFileUploadSource.Initialize", "Content type cannot be empty."
    SetSource SourcePath
    mContentType = ContentType
End Sub

Private Function IHttpUploadSource_ContentLength() As Currency
    IHttpUploadSource_ContentLength = mSize
End Function

Private Function IHttpUploadSource_ContentType() As String
    IHttpUploadSource_ContentType = mContentType
End Function

Private Sub IHttpUploadSource_WriteTo(ByVal Writer As WinHttpUploadStreamWriter)
    Dim reader As New WinHttpUploadFileReader
    Dim buffer() As Byte
    Dim bytesRead As Long
    Dim observedSize As Currency
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo Failed
    reader.OpenExisting mPath
    observedSize = reader.FileSize
    If observedSize <> mSize Then HttpErrors.RaiseTransport HttpErrorIo, "WinHttpNativeTransport.UploadFile", "Upload source file changed during preparation."
    Do While reader.ReadChunk(buffer, bytesRead)
        Writer.WriteBytes buffer, 0, bytesRead
    Loop
    reader.CloseReader
    If Writer.BytesWritten <> mSize Then HttpErrors.RaiseTransport HttpErrorIo, "WinHttpNativeTransport.UploadFile", "Upload source file length changed while reading."
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo CleanupFailed
    reader.CloseReader
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
    Exit Sub

CleanupFailed:
    Err.Clear
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub SetSource(ByVal SourcePath As String)
    mPath = WinHttpUploadFileSystem.NormalizeSource(SourcePath)
    mSize = WinHttpUploadFileSystem.FileSize(mPath)
End Sub
