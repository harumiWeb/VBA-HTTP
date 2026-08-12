Attribute VB_Name = "WinHttpUploadEncoding"
Option Explicit

Private Const TextChunkCharacters As Long = 2048
Private mBoundarySequence As Long

Public Function NextBoundary() As String
    Dim clockPart As Long

    mBoundarySequence = mBoundarySequence + 1
    If mBoundarySequence <= 0 Then mBoundarySequence = 1
    clockPart = CLng((Timer - Fix(Timer)) * 1000000#)
    NextBoundary = "----vba-http-" & Hex$(clockPart) & "-" & Hex$(mBoundarySequence)
End Function

Public Function Utf8ByteCount(ByVal value As String) As Currency
    Dim position As Long
    Dim chunkLength As Long
    Dim bytes As Variant
    Dim total As Currency

    Do While position < Len(value)
        position = position + 1
        chunkLength = NextChunkLength(value, position)
        bytes = HttpEncoding.EncodeUtf8(Mid$(value, position, chunkLength))
        total = AddByteCount(total, HttpEncoding.ByteCount(bytes))
        position = position + chunkLength - 1
    Loop
    Utf8ByteCount = total
End Function

Public Sub WriteUtf8(ByVal value As String, ByVal Writer As WinHttpUploadStreamWriter)
    Dim position As Long
    Dim chunkLength As Long
    Dim bytes As Variant
    Dim buffer() As Byte
    Dim count As Long

    If Writer Is Nothing Then HttpErrors.RaiseValidation "WinHttpUploadEncoding.WriteUtf8", "Writer cannot be Nothing."
    Do While position < Len(value)
        position = position + 1
        chunkLength = NextChunkLength(value, position)
        bytes = HttpEncoding.EncodeUtf8(Mid$(value, position, chunkLength))
        count = HttpEncoding.ByteCount(bytes)
        If count > 0 Then
            buffer = bytes
            Writer.WriteBytes buffer, 0, count
        End If
        position = position + chunkLength - 1
    Loop
End Sub

Public Function ByteCountText(ByVal value As Currency) As String
    If value < 0 Then HttpErrors.RaiseValidation "WinHttpUploadEncoding.ByteCountText", "Byte count cannot be negative."
    ByteCountText = Format$(value, "0")
End Function

Private Function NextChunkLength(ByVal value As String, ByVal position As Long) As Long
    Dim remaining As Long
    Dim length As Long
    Dim code As Long

    remaining = Len(value) - position + 1
    length = remaining
    If length > TextChunkCharacters Then length = TextChunkCharacters
    If position + length - 1 < Len(value) Then
        code = AscW(Mid$(value, position + length - 1, 1))
        If code < 0 Then code = code + 65536
        If code >= 55296 And code <= 56319 Then length = length + 1
    End If
    NextChunkLength = length
End Function

Private Function AddByteCount(ByVal currentValue As Currency, ByVal increment As Long) As Currency
    On Error GoTo Overflow
    AddByteCount = currentValue + CCur(increment)
    Exit Function

Overflow:
    On Error GoTo 0
    HttpErrors.RaiseTransport HttpErrorIo, "WinHttpUploadEncoding", "UTF-8 byte count overflowed the supported range."
End Function
