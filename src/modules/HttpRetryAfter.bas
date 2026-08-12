Attribute VB_Name = "HttpRetryAfter"
Option Explicit

Option Private Module

Public Function TryParse(ByVal value As String, ByVal nowUtc As Date, ByRef delayMilliseconds As Long) As Boolean
    Dim target As Date
    Dim seconds As Double

    value = Trim$(value)
    If Len(value) = 0 Then Exit Function
    If IsAsciiDigits(value) Then
        seconds = ParsePositiveNumber(value)
        delayMilliseconds = SaturatedMilliseconds(seconds)
        TryParse = True
        Exit Function
    End If
    If Not TryParseHttpDate(value, target) Then Exit Function
    seconds = (CDbl(target) - CDbl(nowUtc)) * 86400#
    If seconds < 0 Then seconds = 0
    delayMilliseconds = SaturatedMilliseconds(seconds)
    TryParse = True
End Function

Private Function TryParseHttpDate(ByVal value As String, ByRef result As Date) As Boolean
    Dim normalized As String
    Dim parts As Variant

    On Error GoTo InvalidDate
    normalized = NormalizeSpaces(Replace(value, ",", ""))
    parts = Split(normalized, " ")
    Select Case UBound(parts)
    Case 5
        If UCase$(CStr(parts(5))) = "GMT" Then TryParseHttpDate = ParseImfFixdate(CStr(parts(3)), CStr(parts(2)), CStr(parts(1)), CStr(parts(4)), result)
    Case 3
        If UCase$(CStr(parts(3))) = "GMT" Then TryParseHttpDate = ParseRfc850(CStr(parts(1)), CStr(parts(2)), result)
    Case 4
        TryParseHttpDate = ParseAsctime(CStr(parts(4)), CStr(parts(1)), CStr(parts(2)), CStr(parts(3)), result)
    End Select
    Exit Function

InvalidDate:
    Err.Clear
    TryParseHttpDate = False
End Function

Private Function ParseImfFixdate(ByVal yearText As String, ByVal monthText As String, ByVal dayText As String, ByVal timeText As String, ByRef result As Date) As Boolean
    ParseImfFixdate = BuildDate(CLng(yearText), MonthNumber(monthText), CLng(dayText), timeText, result)
End Function

Private Function ParseRfc850(ByVal dateText As String, ByVal timeText As String, ByRef result As Date) As Boolean
    Dim dateParts As Variant
    Dim yearValue As Long

    dateParts = Split(dateText, "-")
    If UBound(dateParts) <> 2 Then Exit Function
    yearValue = CLng(dateParts(2))
    If yearValue < 100 Then
        If yearValue >= 70 Then yearValue = yearValue + 1900 Else yearValue = yearValue + 2000
    End If
    ParseRfc850 = BuildDate(yearValue, MonthNumber(CStr(dateParts(1))), CLng(dateParts(0)), timeText, result)
End Function

Private Function ParseAsctime(ByVal yearText As String, ByVal monthText As String, ByVal dayText As String, ByVal timeText As String, ByRef result As Date) As Boolean
    ParseAsctime = BuildDate(CLng(yearText), MonthNumber(monthText), CLng(dayText), timeText, result)
End Function

Private Function BuildDate(ByVal yearValue As Long, ByVal monthValue As Long, ByVal dayValue As Long, ByVal timeText As String, ByRef result As Date) As Boolean
    Dim timeParts As Variant

    If monthValue = 0 Then Exit Function
    timeParts = Split(timeText, ":")
    If UBound(timeParts) <> 2 Then Exit Function
    result = DateSerial(yearValue, monthValue, dayValue) + TimeSerial(CLng(timeParts(0)), CLng(timeParts(1)), CLng(timeParts(2)))
    If Year(result) <> yearValue Or Month(result) <> monthValue Or Day(result) <> dayValue Then Exit Function
    If Hour(result) <> CLng(timeParts(0)) Or Minute(result) <> CLng(timeParts(1)) Or Second(result) <> CLng(timeParts(2)) Then Exit Function
    BuildDate = True
End Function

Private Function MonthNumber(ByVal value As String) As Long
    Select Case LCase$(value)
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

Private Function NormalizeSpaces(ByVal value As String) As String
    value = Trim$(value)
    Do While InStr(1, value, "  ", vbBinaryCompare) > 0
        value = Replace(value, "  ", " ")
    Loop
    NormalizeSpaces = value
End Function

Private Function IsAsciiDigits(ByVal value As String) As Boolean
    Dim index As Long
    Dim code As Long

    For index = 1 To Len(value)
        code = AscW(Mid$(value, index, 1))
        If code < 48 Or code > 57 Then Exit Function
    Next index
    IsAsciiDigits = True
End Function

Private Function ParsePositiveNumber(ByVal value As String) As Double
    Dim index As Long

    For index = 1 To Len(value)
        ParsePositiveNumber = ParsePositiveNumber * 10# + (AscW(Mid$(value, index, 1)) - 48)
        If ParsePositiveNumber >= 2147483.647 Then
            ParsePositiveNumber = 2147483.647
            Exit Function
        End If
    Next index
End Function

Private Function SaturatedMilliseconds(ByVal seconds As Double) As Long
    If seconds <= 0 Then Exit Function
    If seconds >= 2147483.647 Then
        SaturatedMilliseconds = 2147483647
    Else
        SaturatedMilliseconds = CLng(Fix(seconds * 1000# + 0.5))
    End If
End Function
