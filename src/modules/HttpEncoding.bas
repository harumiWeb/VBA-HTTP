Attribute VB_Name = "HttpEncoding"
Option Explicit

Private Const Base64Alphabet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Const UnicodeBmpMax As Long = 65535
Private Const UnicodeSupplementaryBase As Long = 65536
Private Const Utf16HighSurrogateMin As Long = 55296
Private Const Utf16HighSurrogateMax As Long = 56319
Private Const Utf16LowSurrogateMin As Long = 56320
Private Const Utf16LowSurrogateMax As Long = 57343
Private Const Utf16PayloadRange As Long = 1024
Private Const Utf8ContinuationPayloadMask As Long = 63
Private Const Utf8ContinuationMarker As Long = 128
Private Const Utf8ContinuationMarkerMask As Long = 192
Private Const Utf8TwoByteMarker As Long = 192
Private Const Utf8ThreeByteMarker As Long = 224
Private Const Utf8FourByteMarker As Long = 240

Public Function PercentEncode(ByVal value As String) As String
    Dim bytes As Variant
    Dim index As Long
    Dim currentByte As Long
    Dim result As String

    bytes = EncodeUtf8(value)
    If Not HasByteArray(bytes) Then
        PercentEncode = ""
        Exit Function
    End If

    ' HasByteArray proved that the Variant contains an allocated array.
    For index = LBound(bytes) To UBound(bytes) ' xlflow:disable-line VBA227
        currentByte = CLng(bytes(index)) ' xlflow:disable-line VBA227
        If IsUnreservedByte(currentByte) Then
            result = result & Chr$(currentByte)
        Else
            result = result & "%" & Right$("0" & Hex$(currentByte), 2)
        End If
    Next index
    PercentEncode = result
End Function

Public Function EncodeUtf8(ByVal value As String) As Variant
    Dim output() As Byte
    Dim byteCount As Long
    Dim index As Long
    Dim codePoint As Long
    Dim firstUnit As Long
    Dim secondUnit As Long

    If Len(value) = 0 Then
        EncodeUtf8 = Array()
        Exit Function
    End If

    ReDim output(0 To Len(value) * 4 - 1)
    index = 1
    Do While index <= Len(value)
        firstUnit = AscW(Mid$(value, index, 1))
        If firstUnit < 0 Then firstUnit = firstUnit + 65536

        If firstUnit >= Utf16HighSurrogateMin And firstUnit <= Utf16HighSurrogateMax Then
            If index = Len(value) Then
                HttpErrors.RaiseValidation "HttpEncoding.EncodeUtf8", "Text contains an unmatched UTF-16 high surrogate."
            End If
            secondUnit = AscW(Mid$(value, index + 1, 1))
            If secondUnit < 0 Then secondUnit = secondUnit + 65536
            If secondUnit < Utf16LowSurrogateMin Or secondUnit > Utf16LowSurrogateMax Then
                HttpErrors.RaiseValidation "HttpEncoding.EncodeUtf8", "Text contains an unmatched UTF-16 high surrogate."
            End If
            codePoint = UnicodeSupplementaryBase + (firstUnit - Utf16HighSurrogateMin) * Utf16PayloadRange + (secondUnit - Utf16LowSurrogateMin)
            index = index + 2
        ElseIf firstUnit >= Utf16LowSurrogateMin And firstUnit <= Utf16LowSurrogateMax Then
            HttpErrors.RaiseValidation "HttpEncoding.EncodeUtf8", "Text contains an unmatched UTF-16 low surrogate."
        Else
            codePoint = firstUnit
            index = index + 1
        End If

        AppendCodePoint output, byteCount, codePoint
    Loop

    ReDim Preserve output(0 To byteCount - 1)
    EncodeUtf8 = output
End Function

Public Function EncodeBase64(ByVal value As Variant) As String
    Dim index As Long
    Dim upperBound As Long
    Dim firstByte As Long
    Dim secondByte As Long
    Dim thirdByte As Long
    Dim byteCount As Long
    Dim result As String

    If Not HasByteArray(value) Then Exit Function
    index = LBound(value) ' xlflow:disable-line VBA227
    upperBound = UBound(value) ' xlflow:disable-line VBA227
    Do While index <= upperBound ' xlflow:disable-line VBA227
        firstByte = CLng(value(index)) ' xlflow:disable-line VBA227
        index = index + 1
        byteCount = 1
        secondByte = 0
        thirdByte = 0
        If index <= upperBound Then
            secondByte = CLng(value(index)) ' xlflow:disable-line VBA227
            index = index + 1
            byteCount = 2
        End If
        If index <= upperBound Then
            thirdByte = CLng(value(index)) ' xlflow:disable-line VBA227
            index = index + 1
            byteCount = 3
        End If

        result = result & Base64Character(firstByte \ 4)
        result = result & Base64Character(((firstByte And 3) * 16) Or (secondByte \ 16))
        If byteCount >= 2 Then
            result = result & Base64Character(((secondByte And 15) * 4) Or (thirdByte \ 64))
        Else
            result = result & "="
        End If
        If byteCount = 3 Then
            result = result & Base64Character(thirdByte And 63)
        Else
            result = result & "="
        End If
    Loop
    EncodeBase64 = result
End Function

Public Function DecodeUtf8(ByVal bytes As Variant) As String
    Dim index As Long
    Dim upperBound As Long
    Dim firstByte As Long
    Dim codePoint As Long
    Dim result As String

    If Not HasByteArray(bytes) Then Exit Function
    ' HasByteArray proved that the Variant contains an allocated array.
    index = LBound(bytes) ' xlflow:disable-line VBA227
    upperBound = UBound(bytes) ' xlflow:disable-line VBA227

    Do While index <= upperBound ' xlflow:disable-line VBA227
        firstByte = CLng(bytes(index)) ' xlflow:disable-line VBA227
        If firstByte <= 127 Then
            codePoint = firstByte
            index = index + 1
        ElseIf firstByte >= 194 And firstByte <= 223 Then
            RequireContinuation bytes, index, upperBound, 1
            codePoint = (firstByte And 31) * 64 + (CLng(bytes(index + 1)) And Utf8ContinuationPayloadMask) ' xlflow:disable-line VBA227
            index = index + 2
        ElseIf firstByte >= 224 And firstByte <= 239 Then
            RequireContinuation bytes, index, upperBound, 2
            If firstByte = 224 And CLng(bytes(index + 1)) < 160 Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
            If firstByte = 237 And CLng(bytes(index + 1)) >= 160 Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
            codePoint = (firstByte And 15) * 4096 + (CLng(bytes(index + 1)) And Utf8ContinuationPayloadMask) * 64 + (CLng(bytes(index + 2)) And Utf8ContinuationPayloadMask) ' xlflow:disable-line VBA227
            index = index + 3
        ElseIf firstByte >= 240 And firstByte <= 244 Then
            RequireContinuation bytes, index, upperBound, 3
            If firstByte = 240 And CLng(bytes(index + 1)) < 144 Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
            If firstByte = 244 And CLng(bytes(index + 1)) > 143 Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
            codePoint = (firstByte And 7) * 262144 + (CLng(bytes(index + 1)) And Utf8ContinuationPayloadMask) * 4096 + (CLng(bytes(index + 2)) And Utf8ContinuationPayloadMask) * 64 + (CLng(bytes(index + 3)) And Utf8ContinuationPayloadMask) ' xlflow:disable-line VBA227
            index = index + 4
        Else
            RaiseMalformedUtf8
        End If
        result = result & CodePointToString(codePoint)
    Loop
    DecodeUtf8 = result
End Function

Public Function DecodeAscii(ByVal bytes As Variant) As String
    Dim index As Long
    Dim result As String

    If Not HasByteArray(bytes) Then Exit Function
    ' HasByteArray proved that the Variant contains an allocated array.
    For index = LBound(bytes) To UBound(bytes) ' xlflow:disable-line VBA227
        If CLng(bytes(index)) > 127 Then ' xlflow:disable-line VBA227
            HttpErrors.RaiseValidation "HttpEncoding.DecodeAscii", "US-ASCII response contains a byte above 127."
        End If
        result = result & Chr$(CLng(bytes(index))) ' xlflow:disable-line VBA227
    Next index
    DecodeAscii = result
End Function

Public Function HasByteArray(ByVal value As Variant) As Boolean
    Dim lowerBound As Long

    If Not IsArray(value) Then Exit Function
    On Error GoTo EmptyArray
    ' Error 9 is the defined empty-array representation returned by Array().
    lowerBound = LBound(value) ' xlflow:disable-line VBA227
    HasByteArray = (UBound(value) >= lowerBound) ' xlflow:disable-line VBA227
    Exit Function
EmptyArray:
    Err.Clear
    HasByteArray = False
    On Error GoTo 0
End Function

Public Function CopyBytes(ByVal value As Variant) As Variant
    Dim output() As Byte
    Dim sourceIndex As Long
    Dim targetIndex As Long

    If Not HasByteArray(value) Then
        CopyBytes = Array()
        Exit Function
    End If
    If VarType(value) = (vbArray Or vbByte) Then
        ' Variant assignment performs a defensive SAFEARRAY copy without a VBA per-byte loop.
        CopyBytes = value
        Exit Function
    End If
    ' HasByteArray proved that the Variant contains an allocated array.
    ReDim output(0 To UBound(value) - LBound(value))
    For sourceIndex = LBound(value) To UBound(value) ' xlflow:disable-line VBA227
        output(targetIndex) = CByte(value(sourceIndex)) ' xlflow:disable-line VBA227
        targetIndex = targetIndex + 1
    Next sourceIndex
    CopyBytes = output
End Function

Private Sub AppendCodePoint(ByRef output() As Byte, ByRef byteCount As Long, ByVal codePoint As Long)
    If codePoint <= 127 Then
        output(byteCount) = CByte(codePoint)
        byteCount = byteCount + 1
    ElseIf codePoint <= 2047 Then
        output(byteCount) = CByte(Utf8TwoByteMarker Or (codePoint \ 64))
        output(byteCount + 1) = CByte(Utf8ContinuationMarker Or (codePoint And Utf8ContinuationPayloadMask))
        byteCount = byteCount + 2
    ElseIf codePoint <= UnicodeBmpMax Then
        output(byteCount) = CByte(Utf8ThreeByteMarker Or (codePoint \ 4096))
        output(byteCount + 1) = CByte(Utf8ContinuationMarker Or ((codePoint \ 64) And Utf8ContinuationPayloadMask))
        output(byteCount + 2) = CByte(Utf8ContinuationMarker Or (codePoint And Utf8ContinuationPayloadMask))
        byteCount = byteCount + 3
    Else
        output(byteCount) = CByte(Utf8FourByteMarker Or (codePoint \ 262144))
        output(byteCount + 1) = CByte(Utf8ContinuationMarker Or ((codePoint \ 4096) And Utf8ContinuationPayloadMask))
        output(byteCount + 2) = CByte(Utf8ContinuationMarker Or ((codePoint \ 64) And Utf8ContinuationPayloadMask))
        output(byteCount + 3) = CByte(Utf8ContinuationMarker Or (codePoint And Utf8ContinuationPayloadMask))
        byteCount = byteCount + 4
    End If
End Sub

Private Sub RequireContinuation(ByVal bytes As Variant, ByVal index As Long, ByVal upperBound As Long, ByVal count As Long)
    Dim offset As Long

    If index + count > upperBound Then RaiseMalformedUtf8
    For offset = 1 To count ' xlflow:disable-line VBA227
        ' DecodeUtf8 proved bounds and RequireContinuation proved this offset.
        If (CLng(bytes(index + offset)) And Utf8ContinuationMarkerMask) <> Utf8ContinuationMarker Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
    Next offset
End Sub

Public Function ByteCount(ByVal value As Variant) As Long
    If Not HasByteArray(value) Then Exit Function
    ByteCount = UBound(value) - LBound(value) + 1 ' xlflow:disable-line VBA227
End Function

Public Function ByteAt(ByVal value As Variant, ByVal index As Long) As Long
    If Not HasByteArray(value) Then
        HttpErrors.RaiseValidation "HttpEncoding.ByteAt", "Byte array is empty."
    End If
    If index < 0 Or index >= ByteCount(value) Then
        HttpErrors.RaiseValidation "HttpEncoding.ByteAt", "Byte index is out of range."
    End If
    ByteAt = CLng(value(LBound(value) + index)) ' xlflow:disable-line VBA227
End Function

Public Sub SetByteAt(ByRef value As Variant, ByVal index As Long, ByVal newValue As Byte)
    If Not HasByteArray(value) Then
        HttpErrors.RaiseValidation "HttpEncoding.SetByteAt", "Byte array is empty."
    End If
    If index < 0 Or index >= ByteCount(value) Then
        HttpErrors.RaiseValidation "HttpEncoding.SetByteAt", "Byte index is out of range."
    End If
    value(LBound(value) + index) = newValue ' xlflow:disable-line VBA227
End Sub

Private Sub RaiseMalformedUtf8()
    HttpErrors.RaiseValidation "HttpEncoding.DecodeUtf8", "Response contains malformed UTF-8."
End Sub

Private Function CodePointToString(ByVal codePoint As Long) As String
    Dim adjusted As Long

    If codePoint <= UnicodeBmpMax Then
        CodePointToString = ChrW$(codePoint)
    Else
        adjusted = codePoint - UnicodeSupplementaryBase
        CodePointToString = ChrW$(Utf16HighSurrogateMin Or (adjusted \ Utf16PayloadRange)) & ChrW$(Utf16LowSurrogateMin Or (adjusted And 1023))
    End If
End Function

Private Function IsUnreservedByte(ByVal value As Long) As Boolean
    IsUnreservedByte = (value >= 65 And value <= 90) Or _
        (value >= 97 And value <= 122) Or _
        (value >= 48 And value <= 57) Or _
        value = 45 Or value = 46 Or value = 95 Or value = 126
End Function

Private Function Base64Character(ByVal index As Long) As String
    Base64Character = Mid$(Base64Alphabet, index + 1, 1)
End Function
