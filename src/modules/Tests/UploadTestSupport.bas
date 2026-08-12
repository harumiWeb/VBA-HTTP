Attribute VB_Name = "UploadTestSupport"
Option Explicit

Public Function NewUploadSource(ByVal Suffix As String) As String
    Dim fileSystem As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    NewUploadSource = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), "vba-http-upload-test-" & fileSystem.GetTempName & "-" & Suffix & ".bin")
End Function

Public Sub WriteUploadPattern(ByVal FilePath As String, ByVal ByteCount As Long)
    Dim fileNumber As Integer
    Dim buffer(0 To 65535) As Byte
    Dim index As Long
    Dim remaining As Long
    Dim requested As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If ByteCount < 0 Then HttpErrors.RaiseValidation "UploadTestSupport.WriteUploadPattern", "Byte count cannot be negative."
    For index = 0 To UBound(buffer)
        buffer(index) = CByte(index Mod 251)
    Next index
    On Error GoTo Cleanup
    fileNumber = FreeFile
    Open FilePath For Binary Access Write Lock Read Write As #fileNumber ' xlflow:disable-line VBA219
    remaining = ByteCount
    Do While remaining > 0
        requested = remaining
        If requested > UBound(buffer) + 1 Then requested = UBound(buffer) + 1
        If requested = UBound(buffer) + 1 Then
            Put #fileNumber, , buffer
        Else
            Put #fileNumber, , LeftBytes(buffer, requested)
        End If
        remaining = remaining - requested
    Loop
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

Public Sub WriteUploadRepeatedByte(ByVal FilePath As String, ByVal ByteValue As Byte, ByVal ByteCount As Long)
    Dim fileNumber As Integer
    Dim buffer(0 To 65535) As Byte
    Dim index As Long
    Dim remaining As Long
    Dim requested As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If ByteCount < 0 Then HttpErrors.RaiseValidation "UploadTestSupport.WriteUploadRepeatedByte", "Byte count cannot be negative."
    For index = 0 To UBound(buffer)
        buffer(index) = ByteValue
    Next index
    On Error GoTo Cleanup
    fileNumber = FreeFile
    Open FilePath For Binary Access Write Lock Read Write As #fileNumber ' xlflow:disable-line VBA219
    remaining = ByteCount
    Do While remaining > 0
        requested = remaining
        If requested > UBound(buffer) + 1 Then requested = UBound(buffer) + 1
        If requested = UBound(buffer) + 1 Then
            Put #fileNumber, , buffer
        Else
            Put #fileNumber, , LeftBytes(buffer, requested)
        End If
        remaining = remaining - requested
    Loop
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

Public Sub DeleteUploadFile(ByVal FilePath As String)
    If Len(FilePath) > 0 Then WinHttpDownloadFileSystem.DeleteFileIfExists FilePath
End Sub

Private Function LeftBytes(ByRef Source() As Byte, ByVal Count As Long) As Variant
    Dim output() As Byte
    Dim index As Long

    ReDim output(0 To Count - 1)
    For index = 0 To Count - 1
        output(index) = Source(index)
    Next index
    LeftBytes = output
End Function
