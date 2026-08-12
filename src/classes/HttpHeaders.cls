VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mNames As New Collection
Private mValues As New Collection

Public Property Get Count() As Long
    Count = mNames.Count
End Property

Public Sub Add(ByVal Name As String, ByVal value As String)
    ValidateName Name
    ValidateValue value
    mNames.Add Name
    mValues.Add value
End Sub

Public Sub SetValue(ByVal Name As String, ByVal value As String)
    ValidateName Name
    ValidateValue value
    Remove Name
    Add Name, value
End Sub

Public Function GetValue(ByVal Name As String) As String
    Dim index As Long

    For index = 1 To mNames.Count
        If NamesEqual(CStr(mNames(index)), Name) Then
            GetValue = CStr(mValues(index))
            Exit Function
        End If
    Next index
End Function

Public Function GetValues(ByVal Name As String) As Variant
    Dim output() As String
    Dim matchCount As Long
    Dim index As Long

    For index = 1 To mNames.Count
        If NamesEqual(CStr(mNames(index)), Name) Then matchCount = matchCount + 1
    Next index

    If matchCount = 0 Then
        GetValues = Array()
    Else
        ReDim output(0 To matchCount - 1)
        matchCount = 0
        For index = 1 To mNames.Count
            If NamesEqual(CStr(mNames(index)), Name) Then
                output(matchCount) = CStr(mValues(index))
                matchCount = matchCount + 1
            End If
        Next index
        GetValues = output
    End If
End Function

Public Function Contains(ByVal Name As String) As Boolean
    Dim index As Long

    For index = 1 To mNames.Count
        If NamesEqual(CStr(mNames(index)), Name) Then
            Contains = True
            Exit Function
        End If
    Next index
End Function

Public Sub Remove(ByVal Name As String)
    Dim index As Long

    index = mNames.Count
    Do While index >= 1
        If NamesEqual(CStr(mNames(index)), Name) Then
            mNames.Remove index
            mValues.Remove index
        End If
        index = index - 1
    Loop
End Sub

Public Sub Clear()
    Set mNames = New Collection
    Set mValues = New Collection
End Sub

Public Function NameAt(ByVal index As Long) As String
    NameAt = CStr(mNames(index))
End Function

Public Function ValueAt(ByVal index As Long) As String
    ValueAt = CStr(mValues(index))
End Function

Public Function Clone() As HttpHeaders
    Dim copy As New HttpHeaders
    Dim index As Long

    For index = 1 To mNames.Count
        copy.Add CStr(mNames(index)), CStr(mValues(index))
    Next index
    Set Clone = copy
End Function

Private Sub Class_Initialize()
    Clear
End Sub

Private Sub ValidateName(ByVal Name As String)
    Dim index As Long
    Dim code As Long

    If Len(Name) = 0 Or Trim$(Name) <> Name Then
        HttpErrors.RaiseValidation "HttpHeaders", "Header name must be non-empty and cannot contain surrounding whitespace."
    End If
    For index = 1 To Len(Name)
        code = AscW(Mid$(Name, index, 1))
        If code <= 32 Or code >= 127 Or code = 58 Then
            HttpErrors.RaiseValidation "HttpHeaders", "Header name contains an invalid character."
        End If
    Next index
End Sub

Private Sub ValidateValue(ByVal value As String)
    If InStr(1, value, vbCr, vbBinaryCompare) > 0 Or InStr(1, value, vbLf, vbBinaryCompare) > 0 Then
        HttpErrors.RaiseValidation "HttpHeaders", "Header value cannot contain CR or LF."
    End If
End Sub

Private Function NamesEqual(ByVal first As String, ByVal second As String) As Boolean
    NamesEqual = (StrComp(first, second, vbTextCompare) = 0)
End Function
