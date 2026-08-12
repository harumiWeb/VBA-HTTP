VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpDownloadFileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFileNumber As Integer
Private mPath As String
Private mIsOpen As Boolean

Friend Sub OpenNew(ByVal FilePath As String)
    If Len(FilePath) = 0 Then HttpErrors.RaiseValidation "WinHttpDownloadFileWriter.OpenNew", "Temporary file path cannot be empty."
    If mIsOpen Then HttpErrors.RaiseValidation "WinHttpDownloadFileWriter.OpenNew", "Writer is already open."

    mPath = FilePath
    mFileNumber = FreeFile
    Open mPath For Binary Access Write Lock Read Write As #mFileNumber
    mIsOpen = True
End Sub

Friend Sub WriteChunk(ByRef Buffer() As Byte)
    If Not mIsOpen Then HttpErrors.RaiseValidation "WinHttpDownloadFileWriter.WriteChunk", "Writer is not open."
    Put #mFileNumber, , Buffer
End Sub

Friend Sub FlushAndClose()
    If Not mIsOpen Then Exit Sub
    Close #mFileNumber
    mIsOpen = False
End Sub

Friend Function Abort() As Boolean
    Abort = False
    On Error GoTo CleanupFailed
    If mIsOpen Then Close #mFileNumber
    mIsOpen = False
    If Len(mPath) > 0 Then WinHttpDownloadFileSystem.DeleteFileIfExists mPath
    Abort = True
    On Error GoTo 0
    Exit Function

CleanupFailed:
    Err.Clear
    On Error GoTo 0
End Function

Friend Property Get IsOpen() As Boolean
    IsOpen = mIsOpen
End Property

Friend Property Get FilePath() As String
    FilePath = mPath
End Property

Private Sub Class_Terminate()
    If mIsOpen Then
        If Not Abort() Then Err.Clear
    End If
End Sub
