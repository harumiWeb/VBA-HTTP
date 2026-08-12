Attribute VB_Name = "HttpJson"
Option Explicit

Public Function Quote(ByVal value As String) As String
    Dim index As Long
    Dim code As Long
    Dim character As String
    Dim result As String

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        code = AscW(character)
        If code < 0 Then code = code + 65536
        Select Case code
        Case 8
            result = result & "\b"
        Case 9
            result = result & "\t"
        Case 10
            result = result & "\n"
        Case 12
            result = result & "\f"
        Case 13
            result = result & "\r"
        Case 34
            result = result & Chr$(92) & Chr$(34)
        Case 92
            result = result & Chr$(92) & Chr$(92)
        Case 0 To 31
            result = result & "\u" & Right$("0000" & Hex$(code), 4)
        Case Else
            result = result & character
        End Select
    Next index
    Quote = Chr$(34) & result & Chr$(34)
End Function

Public Function Number(ByVal value As Double) As String
    Number = Trim$(Str$(Round(value, 3)))
    Number = Replace$(Number, ",", ".")
    If Left$(Number, 1) = "." Then
        Number = "0" & Number
    ElseIf Left$(Number, 2) = "-." Then
        Number = "-0" & Mid$(Number, 2)
    End If
End Function

Public Function BooleanValue(ByVal value As Boolean) As String
    If value Then
        BooleanValue = "true"
    Else
        BooleanValue = "false"
    End If
End Function

Public Function HeadersArray(ByVal Headers As HttpHeaders) As String
    Dim index As Long
    Dim result As String

    result = "["
    If Not Headers Is Nothing Then
        For index = 1 To Headers.Count
            If index > 1 Then result = result & ","
            result = result & "{" & Chr$(34) & "name" & Chr$(34) & ":" & Quote(Headers.NameAt(index)) & "," & _
                Chr$(34) & "value" & Chr$(34) & ":" & Quote(Headers.ValueAt(index)) & "}"
        Next index
    End If
    HeadersArray = result & "]"
End Function
