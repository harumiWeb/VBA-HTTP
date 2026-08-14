VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpUploadFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const UploadChunkBytes As Long = 65536
Private mFileNumber As Integer
Private mPath As String
Private mSize As Currency
Private mPosition As Currency
Private mIsOpen As Boolean
Private mBufferCapacity As Long

Friend Sub OpenExisting(ByVal FilePath As String)
    If Len(FilePath) = 0 Then HttpErrors.RaiseValidation "WinHttpUploadFileReader.OpenExisting", "Source file path cannot be empty."
    If mIsOpen Then HttpErrors.RaiseValidation "WinHttpUploadFileReader.OpenExisting", "Reader is already open."

    mPath = FilePath
    mSize = WinHttpUploadFileSystem.FileSize(mPath)
    mPosition = 0
    mBufferCapacity = 0
    mFileNumber = FreeFile
    Open mPath For Binary Access Read Lock Read Write As #mFileNumber
    mIsOpen = True
End Sub

Friend Function ReadChunk(ByRef Buffer() As Byte, ByRef BytesRead As Long) As Boolean
    Dim remaining As Currency
    Dim requested As Long

    If Not mIsOpen Then HttpErrors.RaiseValidation "WinHttpUploadFileReader.ReadChunk", "Reader is not open."
    remaining = mSize - mPosition
    If remaining <= 0 Then Exit Function
    If remaining > UploadChunkBytes Then
        requested = UploadChunkBytes
    Else
        requested = CLng(remaining)
    End If
    If requested <> mBufferCapacity Then
        ReDim Buffer(0 To requested - 1)
        mBufferCapacity = requested
    End If
    Get #mFileNumber, , Buffer
    mPosition = mPosition + CCur(requested)
    BytesRead = requested
    ReadChunk = True
End Function

Friend Property Get FileSize() As Currency
    FileSize = mSize
End Property

Friend Sub CloseReader()
    If Not mIsOpen Then Exit Sub
    Close #mFileNumber
    mFileNumber = 0
    mIsOpen = False
End Sub

Private Sub Class_Terminate()
    If mIsOpen Then
        If Not CloseReaderSafely() Then Err.Clear
    End If
End Sub

Private Function CloseReaderSafely() As Boolean
    On Error GoTo Failed
    Close #mFileNumber
    mFileNumber = 0
    mIsOpen = False
    CloseReaderSafely = True
    On Error GoTo 0
    Exit Function

Failed:
    Err.Clear
    CloseReaderSafely = False
    On Error GoTo 0
End Function
