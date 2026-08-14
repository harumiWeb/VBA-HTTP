Attribute VB_Name = "DownloadTestSupport"
Option Explicit

Public Function NewDownloadDestination(ByVal Suffix As String) As String
    Dim fileSystem As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    NewDownloadDestination = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), "vba-http-download-test-" & fileSystem.GetTempName & "-" & Suffix & ".bin")
End Function

Public Sub WriteDownloadSentinel(ByVal FilePath As String)
    Dim fileNumber As Integer
    Dim bytes(0 To 7) As Byte
    Dim index As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    For index = 0 To UBound(bytes)
        bytes(index) = 239 - index
    Next index
    On Error GoTo Cleanup
    fileNumber = FreeFile
    Open FilePath For Binary Access Write Lock Read Write As #fileNumber ' xlflow:disable-line VBA219
    Put #fileNumber, , bytes
    Close #fileNumber
    fileNumber = 0
    Exit Sub

Cleanup:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    Close #fileNumber
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Public Function ReadDownloadByte(ByVal FilePath As String, ByVal Offset As Long) As Byte
    Dim fileNumber As Integer
    Dim value As Byte
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo Cleanup
    fileNumber = FreeFile
    Open FilePath For Binary Access Read Lock Read Write As #fileNumber ' xlflow:disable-line VBA219
    Get #fileNumber, Offset + 1, value
    Close #fileNumber
    fileNumber = 0
    ReadDownloadByte = value
    Exit Function

Cleanup:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    Close #fileNumber
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Public Function CountDownloadTemporaryFiles(ByVal DestinationPath As String) As Long
    Dim fileSystem As Object
    Dim parentFolder As Object
    Dim fileItem As Object
    Dim fileIndex As Long

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set parentFolder = fileSystem.GetFolder(fileSystem.GetParentFolderName(DestinationPath))
    If parentFolder Is Nothing Then Exit Function
    For fileIndex = 1 To parentFolder.Files.Count
        Set fileItem = parentFolder.Files.Item(fileIndex)
        If Not fileItem Is Nothing Then
            If LCase$(Left$(fileItem.Name, Len(".vba-http-download-"))) = ".vba-http-download-" Then CountDownloadTemporaryFiles = CountDownloadTemporaryFiles + 1
        End If
    Next fileIndex
End Function

Public Sub DeleteDownloadFile(ByVal FilePath As String)
    If Len(FilePath) > 0 Then WinHttpDownloadFileSystem.DeleteFileIfExists FilePath
End Sub
