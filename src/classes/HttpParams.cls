VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpParams"
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

Public Sub Add(ByVal Name As String, ByVal value As Variant)
    If Len(Name) = 0 Then
        HttpErrors.RaiseValidation "HttpParams.Add", "Query parameter name cannot be empty."
    End If
    If IsObject(value) Or IsArray(value) Or IsError(value) Or IsNull(value) Then
        HttpErrors.RaiseValidation "HttpParams.Add", "Query parameter value must be a scalar non-Null value."
    End If
    mNames.Add Name
    mValues.Add ScalarText(value)
End Sub

Public Sub SetValue(ByVal Name As String, ByVal value As Variant)
    Remove Name
    Add Name, value
End Sub

Public Sub Remove(ByVal Name As String)
    Dim index As Long

    index = mNames.Count
    Do While index >= 1
        If StrComp(CStr(mNames(index)), Name, vbBinaryCompare) = 0 Then
            mNames.Remove index
            mValues.Remove index
        End If
        index = index - 1
    Loop
End Sub

Public Function ToQueryString() As String
    Dim result As String
    Dim index As Long

    For index = 1 To mNames.Count
        If index > 1 Then result = result & "&"
        result = result & HttpEncoding.PercentEncode(CStr(mNames(index))) & "=" & HttpEncoding.PercentEncode(CStr(mValues(index)))
    Next index
    ToQueryString = result
End Function

Public Function Clone() As HttpParams
    Dim copy As New HttpParams
    Dim index As Long

    For index = 1 To mNames.Count
        copy.Add CStr(mNames(index)), CStr(mValues(index))
    Next index
    Set Clone = copy
End Function

Private Function ScalarText(ByVal value As Variant) As String
    If VarType(value) = vbBoolean Then
        ScalarText = LCase$(CStr(CBool(value)))
    ElseIf VarType(value) = vbDate Then
        ScalarText = Format$(CDate(value), "yyyy-mm-dd\Thh:nn:ss")
    ElseIf IsNumeric(value) Then
        ScalarText = Trim$(Str$(CDbl(value)))
    Else
        ScalarText = CStr(value)
    End If
End Function
