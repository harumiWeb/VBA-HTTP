Attribute VB_Name = "HttpEncoding"
Option Explicit

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

        If firstUnit >= &HD800& And firstUnit <= &HDBFF& Then
            If index = Len(value) Then
                HttpErrors.RaiseValidation "HttpEncoding.EncodeUtf8", "Text contains an unmatched UTF-16 high surrogate."
            End If
            secondUnit = AscW(Mid$(value, index + 1, 1))
            If secondUnit < 0 Then secondUnit = secondUnit + 65536
            If secondUnit < &HDC00& Or secondUnit > &HDFFF& Then
                HttpErrors.RaiseValidation "HttpEncoding.EncodeUtf8", "Text contains an unmatched UTF-16 high surrogate."
            End If
            codePoint = &H10000 + (firstUnit - &HD800&) * &H400 + (secondUnit - &HDC00&)
            index = index + 2
        ElseIf firstUnit >= &HDC00& And firstUnit <= &HDFFF& Then
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
        If firstByte <= &H7F Then
            codePoint = firstByte
            index = index + 1
        ElseIf firstByte >= &HC2 And firstByte <= &HDF Then
            RequireContinuation bytes, index, upperBound, 1
            codePoint = (firstByte And &H1F) * &H40 + (CLng(bytes(index + 1)) And &H3F) ' xlflow:disable-line VBA227
            index = index + 2
        ElseIf firstByte >= &HE0 And firstByte <= &HEF Then
            RequireContinuation bytes, index, upperBound, 2
            If firstByte = &HE0 And CLng(bytes(index + 1)) < &HA0 Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
            If firstByte = &HED And CLng(bytes(index + 1)) >= &HA0 Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
            codePoint = (firstByte And &HF) * &H1000 + (CLng(bytes(index + 1)) And &H3F) * &H40 + (CLng(bytes(index + 2)) And &H3F) ' xlflow:disable-line VBA227
            index = index + 3
        ElseIf firstByte >= &HF0 And firstByte <= &HF4 Then
            RequireContinuation bytes, index, upperBound, 3
            If firstByte = &HF0 And CLng(bytes(index + 1)) < &H90 Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
            If firstByte = &HF4 And CLng(bytes(index + 1)) > &H8F Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
            codePoint = (firstByte And &H7) * &H40000 + (CLng(bytes(index + 1)) And &H3F) * &H1000 + (CLng(bytes(index + 2)) And &H3F) * &H40 + (CLng(bytes(index + 3)) And &H3F) ' xlflow:disable-line VBA227
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
    ' HasByteArray proved that the Variant contains an allocated array.
    ReDim output(0 To UBound(value) - LBound(value))
    For sourceIndex = LBound(value) To UBound(value) ' xlflow:disable-line VBA227
        output(targetIndex) = CByte(value(sourceIndex)) ' xlflow:disable-line VBA227
        targetIndex = targetIndex + 1
    Next sourceIndex
    CopyBytes = output
End Function

Private Sub AppendCodePoint(ByRef output() As Byte, ByRef byteCount As Long, ByVal codePoint As Long)
    If codePoint <= &H7F Then
        output(byteCount) = CByte(codePoint)
        byteCount = byteCount + 1
    ElseIf codePoint <= &H7FF Then
        output(byteCount) = CByte(&HC0 Or (codePoint \ &H40))
        output(byteCount + 1) = CByte(&H80 Or (codePoint And &H3F))
        byteCount = byteCount + 2
    ElseIf codePoint <= &HFFFF& Then
        output(byteCount) = CByte(&HE0 Or (codePoint \ &H1000))
        output(byteCount + 1) = CByte(&H80 Or ((codePoint \ &H40) And &H3F))
        output(byteCount + 2) = CByte(&H80 Or (codePoint And &H3F))
        byteCount = byteCount + 3
    Else
        output(byteCount) = CByte(&HF0 Or (codePoint \ &H40000))
        output(byteCount + 1) = CByte(&H80 Or ((codePoint \ &H1000) And &H3F))
        output(byteCount + 2) = CByte(&H80 Or ((codePoint \ &H40) And &H3F))
        output(byteCount + 3) = CByte(&H80 Or (codePoint And &H3F))
        byteCount = byteCount + 4
    End If
End Sub

Private Sub RequireContinuation(ByVal bytes As Variant, ByVal index As Long, ByVal upperBound As Long, ByVal count As Long)
    Dim offset As Long

    If index + count > upperBound Then RaiseMalformedUtf8
    For offset = 1 To count ' xlflow:disable-line VBA227
        ' DecodeUtf8 proved bounds and RequireContinuation proved this offset.
        If (CLng(bytes(index + offset)) And &HC0) <> &H80 Then RaiseMalformedUtf8 ' xlflow:disable-line VBA227
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

    If codePoint <= &HFFFF& Then
        CodePointToString = ChrW$(codePoint)
    Else
        adjusted = codePoint - &H10000
        CodePointToString = ChrW$(&HD800& Or (adjusted \ &H400)) & ChrW$(&HDC00& Or (adjusted And &H3FF))
    End If
End Function

Private Function IsUnreservedByte(ByVal value As Long) As Boolean
    IsUnreservedByte = (value >= 65 And value <= 90) Or _
        (value >= 97 And value <= 122) Or _
        (value >= 48 And value <= 57) Or _
        value = 45 Or value = 46 Or value = 95 Or value = 126
End Function
