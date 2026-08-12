Attribute VB_Name = "WinHttpDownloadFileSystem"
Option Explicit

Private Const MoveFileReplaceExisting As Long = 1
Private Const MoveFileWriteThrough As Long = 8

#If VBA7 Then
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal ExistingFileName As LongPtr, ByVal NewFileName As LongPtr, ByVal Flags As Long) As Long
Private Declare PtrSafe Function DeleteFileW Lib "kernel32" (ByVal FileName As LongPtr) As Long
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
#Else
Private Declare Function MoveFileExW Lib "kernel32" (ByVal ExistingFileName As Long, ByVal NewFileName As Long, ByVal Flags As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32" (ByVal FileName As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
#End If

Public Function NormalizeDestination(ByVal DestinationPath As String) As String
    Dim fileSystem As Object
    Dim parentPath As String

    DestinationPath = Trim$(DestinationPath)
    If Len(DestinationPath) = 0 Then HttpErrors.RaiseValidation "WinHttpDownloadFileSystem.NormalizeDestination", "Destination path cannot be empty."
    If InStr(1, DestinationPath, "*", vbBinaryCompare) > 0 Or InStr(1, DestinationPath, "?", vbBinaryCompare) > 0 Then
        HttpErrors.RaiseValidation "WinHttpDownloadFileSystem.NormalizeDestination", "Destination path cannot contain wildcard characters."
    End If

    On Error GoTo InvalidPath
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    NormalizeDestination = fileSystem.GetAbsolutePathName(DestinationPath)
    parentPath = fileSystem.GetParentFolderName(NormalizeDestination)
    If Len(parentPath) = 0 Or Not fileSystem.FolderExists(parentPath) Then
        HttpErrors.RaiseValidation "WinHttpDownloadFileSystem.NormalizeDestination", "Destination parent directory does not exist."
    End If
    If fileSystem.FolderExists(NormalizeDestination) Then
        HttpErrors.RaiseValidation "WinHttpDownloadFileSystem.NormalizeDestination", "Destination path is a directory."
    End If
    On Error GoTo 0
    Exit Function

InvalidPath:
    If HttpErrors.CategoryFromNumber(Err.Number) = HttpErrorValidation Then Err.Raise Err.Number, Err.Source, Err.Description
    Err.Raise HttpErrValidation, "WinHttpDownloadFileSystem.NormalizeDestination", "Destination path is invalid."
End Function

Public Function CreateTemporaryPath(ByVal DestinationPath As String) As String
    Dim fileSystem As Object
    Dim parentPath As String
    Dim candidate As String
    Dim attempt As Long

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    parentPath = fileSystem.GetParentFolderName(DestinationPath)
    If Len(parentPath) = 0 Or Not fileSystem.FolderExists(parentPath) Then
        HttpErrors.RaiseValidation "WinHttpDownloadFileSystem.CreateTemporaryPath", "Destination parent directory does not exist."
    End If

    For attempt = 1 To 20
        candidate = fileSystem.BuildPath(parentPath, ".vba-http-download-" & fileSystem.GetTempName)
        If Not fileSystem.FileExists(candidate) Then
            CreateTemporaryPath = candidate
            Exit Function
        End If
    Next attempt
    HttpErrors.RaiseTransport HttpErrorIo, "WinHttpDownloadFileSystem.CreateTemporaryPath", "Could not allocate a unique temporary download path."
End Function

Public Sub ReplaceFileAtomically(ByVal TemporaryPath As String, ByVal DestinationPath As String)
    Dim result As Long
    Dim errorCode As Long

    If Len(TemporaryPath) = 0 Or Len(DestinationPath) = 0 Then HttpErrors.RaiseValidation "WinHttpDownloadFileSystem.ReplaceFileAtomically", "File paths cannot be empty."
    #If VBA7 Then
    result = MoveFileExW(StrPtr(TemporaryPath), StrPtr(DestinationPath), MoveFileReplaceExisting Or MoveFileWriteThrough)
    #Else
    result = MoveFileExW(StrPtr(TemporaryPath), StrPtr(DestinationPath), MoveFileReplaceExisting Or MoveFileWriteThrough)
    #End If
    If result <> 0 Then Exit Sub

    errorCode = Err.LastDllError
    If errorCode = 0 Then errorCode = GetLastError()
    HttpErrors.RaiseTransport HttpErrorIo, "WinHttpDownloadFileSystem.ReplaceFileAtomically", "Atomic download publication failed (" & CStr(errorCode) & ")."
End Sub

Public Sub DeleteFileIfExists(ByVal FilePath As String)
    If Len(FilePath) = 0 Then Exit Sub
    #If VBA7 Then
    Call DeleteFileW(StrPtr(FilePath))
    #Else
    Call DeleteFileW(StrPtr(FilePath))
    #End If
End Sub
