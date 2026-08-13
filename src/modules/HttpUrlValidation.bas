Attribute VB_Name = "HttpUrlValidation"
Option Explicit

Option Private Module

Public Sub ValidateNoUserInfo(ByVal value As String, ByVal Source As String)
    Dim schemeSeparator As Long
    Dim authorityStart As Long
    Dim authorityEnd As Long
    Dim authority As String
    Dim index As Long

    schemeSeparator = InStr(1, value, "://", vbBinaryCompare)
    If schemeSeparator <= 1 Then Exit Sub

    authorityStart = schemeSeparator + 3
    authorityEnd = Len(value) + 1
    For index = authorityStart To Len(value)
        If InStr(1, "/?#", Mid$(value, index, 1), vbBinaryCompare) > 0 Then
            authorityEnd = index
            Exit For
        End If
    Next index

    authority = Mid$(value, authorityStart, authorityEnd - authorityStart)
    If InStr(1, authority, "@", vbBinaryCompare) > 0 Then
        HttpErrors.RaiseInvalidUrl Source, "URL user-info is not permitted."
    End If
End Sub

Public Function WithoutFragment(ByVal value As String) As String
    Dim fragmentIndex As Long

    fragmentIndex = InStr(1, value, "#", vbBinaryCompare)
    If fragmentIndex > 0 Then
        WithoutFragment = Left$(value, fragmentIndex - 1)
    Else
        WithoutFragment = value
    End If
End Function
