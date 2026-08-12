VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpMultipartEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mForm As HttpMultipartForm

Friend Sub Initialize(ByVal Form As HttpMultipartForm)
    If Form Is Nothing Then HttpErrors.RaiseValidation "WinHttpMultipartEncoder.Initialize", "Multipart form cannot be Nothing."
    Set mForm = Form.Clone()
End Sub

Friend Function CloneForm() As HttpMultipartForm
    Set CloneForm = mForm.Clone()
End Function

Friend Function ContentLength() As Currency
    Dim index As Long
    Dim part As HttpMultipartPart
    Dim total As Currency

    For index = 1 To mForm.Count
        Set part = mForm.PartAt(index)
        total = AddLength(total, WinHttpUploadEncoding.Utf8ByteCount(PartPrefix(part)))
        If part.IsFile Then
            total = AddLength(total, WinHttpUploadFileSystem.FileSize(part.FilePath))
        Else
            total = AddLength(total, WinHttpUploadEncoding.Utf8ByteCount(part.Value))
        End If
        total = AddLength(total, 2)
    Next index
    total = AddLength(total, WinHttpUploadEncoding.Utf8ByteCount("--" & mForm.Boundary & "--" & vbCrLf))
    ContentLength = total
End Function

Friend Function ContentType() As String
    ContentType = "multipart/form-data; boundary=" & mForm.Boundary
End Function

Friend Sub WriteTo(ByVal Writer As WinHttpUploadStreamWriter)
    Dim index As Long
    Dim part As HttpMultipartPart
    Dim reader As WinHttpUploadFileReader
    Dim buffer() As Byte
    Dim bytesRead As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo Failed
    For index = 1 To mForm.Count
        Set part = mForm.PartAt(index)
        WinHttpUploadEncoding.WriteUtf8 PartPrefix(part), Writer
        If part.IsFile Then
            Set reader = New WinHttpUploadFileReader
            reader.OpenExisting part.FilePath
            Do While reader.ReadChunk(buffer, bytesRead)
                Writer.WriteBytes buffer, 0, bytesRead
            Loop
            reader.CloseReader
            Set reader = Nothing
        Else
            WinHttpUploadEncoding.WriteUtf8 part.Value, Writer
        End If
        WinHttpUploadEncoding.WriteUtf8 vbCrLf, Writer
    Next index
    WinHttpUploadEncoding.WriteUtf8 "--" & mForm.Boundary & "--" & vbCrLf, Writer
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    If Not reader Is Nothing Then
        On Error GoTo CleanupFailed
        reader.CloseReader
        On Error GoTo 0
    End If
    Err.Raise errorNumber, errorSource, errorDescription
    Exit Sub

CleanupFailed:
    Err.Clear
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Function PartPrefix(ByVal part As HttpMultipartPart) As String
    Dim header As String

    header = "--" & mForm.Boundary & vbCrLf
    header = header & "Content-Disposition: form-data; name=""" & part.Name & """"
    If part.IsFile Then
        header = header & "; filename=""" & part.FileName & """" & vbCrLf
        header = header & "Content-Type: " & part.ContentType
    End If
    PartPrefix = header & vbCrLf & vbCrLf
End Function

Private Function AddLength(ByVal currentValue As Currency, ByVal increment As Variant) As Currency
    On Error GoTo Overflow
    AddLength = currentValue + CCur(increment)
    Exit Function

Overflow:
    On Error GoTo 0
    HttpErrors.RaiseTransport HttpErrorIo, "WinHttpMultipartEncoder.ContentLength", "Multipart content length overflowed the supported range."
End Function
