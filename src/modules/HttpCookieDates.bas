Attribute VB_Name = "HttpCookieDates"
Option Explicit

Public Function TryParse(ByVal value As String, ByRef output As Date) As Boolean
    Dim tokens As Collection
    Dim token As Variant
    Dim dayText As String
    Dim monthText As String
    Dim yearText As String
    Dim timeText As String
    Dim monthValue As Long
    Dim timeParts As Variant
    Dim yearNumber As Long

    Set tokens = TokensOf(value)
    If tokens.Count < 3 Then Exit Function
    On Error GoTo Invalid

    timeText = "00:00:00"
    For Each token In tokens
        If InStr(1, CStr(token), ":", vbBinaryCompare) > 0 Then
            timeText = CStr(token)
            Exit For
        End If
    Next token
    If InStr(1, timeText, ":", vbBinaryCompare) = 0 Then Exit Function
    timeParts = Split(timeText, ":")
    If UBound(timeParts) < 2 Then Exit Function

    If tokens.Count >= 4 Then
        If IsNumericToken(CStr(tokens(1))) And IsMonthToken(CStr(tokens(2))) Then
            dayText = CStr(tokens(1))
            monthText = CStr(tokens(2))
            yearText = CStr(tokens(3))
        ElseIf IsMonthToken(CStr(tokens(1))) And IsNumericToken(CStr(tokens(2))) Then
            monthText = CStr(tokens(1))
            dayText = CStr(tokens(2))
            yearText = CStr(tokens(4))
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    monthValue = MonthNumber(monthText)
    If monthValue = 0 Or Not IsNumericToken(dayText) Or Not IsNumericToken(yearText) Then Exit Function
    yearNumber = CLng(yearText)
    If yearNumber < 100 Then
        If yearNumber >= 70 Then yearNumber = yearNumber + 1900 Else yearNumber = yearNumber + 2000
    End If
    output = DateSerial(yearNumber, monthValue, CLng(dayText)) + _
        TimeSerial(CLng(timeParts(0)), CLng(timeParts(1)), CLng(timeParts(2)))
    TryParse = True
    Exit Function
Invalid:
    TryParse = False
    Err.Clear
End Function

Private Function TokensOf(ByVal value As String) As Collection
    Dim output As New Collection
    Dim normalized As String
    Dim values As Variant
    Dim index As Long
    Dim token As String

    normalized = Replace(Trim$(value), ",", " ")
    normalized = Replace(normalized, "-", " ")
    Do While InStr(1, normalized, "  ", vbBinaryCompare) > 0
        normalized = Replace(normalized, "  ", " ")
    Loop
    values = Split(normalized, " ")
    For index = LBound(values) To UBound(values)
        token = Trim$(CStr(values(index)))
        If Len(token) > 0 And LCase$(token) <> "gmt" Then output.Add token
    Next index
    Set TokensOf = output
End Function

Private Function IsNumericToken(ByVal value As String) As Boolean
    Dim index As Long
    If Len(value) = 0 Then Exit Function
    For index = 1 To Len(value)
        If Not (Mid$(value, index, 1) Like "#") Then Exit Function
    Next index
    IsNumericToken = True
End Function

Private Function IsMonthToken(ByVal value As String) As Boolean
    IsMonthToken = (MonthNumber(value) > 0)
End Function

Private Function MonthNumber(ByVal value As String) As Long
    Select Case LCase$(Left$(Trim$(value), 3))
    Case "jan": MonthNumber = 1
    Case "feb": MonthNumber = 2
    Case "mar": MonthNumber = 3
    Case "apr": MonthNumber = 4
    Case "may": MonthNumber = 5
    Case "jun": MonthNumber = 6
    Case "jul": MonthNumber = 7
    Case "aug": MonthNumber = 8
    Case "sep": MonthNumber = 9
    Case "oct": MonthNumber = 10
    Case "nov": MonthNumber = 11
    Case "dec": MonthNumber = 12
    End Select
End Function
