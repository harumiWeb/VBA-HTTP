Attribute VB_Name = "WinHttpUploadFileSystem"
Option Explicit

Public Function NormalizeSource(ByVal SourcePath As String) As String
    Dim fileSystem As Object
    Dim absolutePath As String

    SourcePath = Trim$(SourcePath)
    If Len(SourcePath) = 0 Then HttpErrors.RaiseValidation "WinHttpUploadFileSystem.NormalizeSource", "Source path cannot be empty."
    If InStr(1, SourcePath, "*", vbBinaryCompare) > 0 Or InStr(1, SourcePath, "?", vbBinaryCompare) > 0 Then
        HttpErrors.RaiseValidation "WinHttpUploadFileSystem.NormalizeSource", "Source path cannot contain wildcard characters."
    End If

    On Error GoTo InvalidPath
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    absolutePath = fileSystem.GetAbsolutePathName(SourcePath)
    If Not fileSystem.FileExists(absolutePath) Or fileSystem.FolderExists(absolutePath) Then
        HttpErrors.RaiseTransport HttpErrorIo, "WinHttpUploadFileSystem.NormalizeSource", "Upload source file does not exist."
    End If
    NormalizeSource = absolutePath
    On Error GoTo 0
    Exit Function

InvalidPath:
    If HttpErrors.CategoryFromNumber(Err.Number) <> HttpErrorNone Then Err.Raise Err.Number, Err.Source, Err.Description
    Err.Raise HttpErrIo, "WinHttpUploadFileSystem.NormalizeSource", "Upload source path is invalid."
End Function

Public Function FileSize(ByVal SourcePath As String) As Currency
    Dim fileSystem As Object
    Dim value As Variant

    On Error GoTo SizeFailed
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    value = fileSystem.GetFile(SourcePath).Size
    If CDbl(value) < 0 Then GoTo SizeFailed
    FileSize = CCur(value)
    Exit Function

SizeFailed:
    On Error GoTo 0
    Err.Raise HttpErrIo, "WinHttpUploadFileSystem.FileSize", "Could not determine upload source size."
End Function
