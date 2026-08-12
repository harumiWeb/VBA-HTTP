VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mKind As HttpBodyKind
Private mText As String
Private mBytes As Variant

Public Property Get Kind() As HttpBodyKind
    Kind = mKind
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = (mKind = HttpBodyEmpty)
End Property

Public Property Get Text() As String
    If mKind = HttpBodyBytes Then
        HttpErrors.RaiseValidation "HttpBody.Text", "Binary body cannot be read as text without a response charset."
    End If
    Text = mText
End Property

Public Property Let Text(ByVal value As String)
    mKind = HttpBodyText
    mText = value
    mBytes = Array()
End Property

Public Property Get Bytes() As Variant
    If mKind = HttpBodyText Then
        Bytes = HttpEncoding.EncodeUtf8(mText)
    Else
        Bytes = HttpEncoding.CopyBytes(mBytes)
    End If
End Property

Public Sub SetBytes(ByVal value As Variant)
    If Not IsArray(value) Then
        HttpErrors.RaiseValidation "HttpBody.SetBytes", "Binary body must be an array."
    End If
    mKind = HttpBodyBytes
    mText = ""
    mBytes = HttpEncoding.CopyBytes(value)
End Sub

Public Sub Clear()
    mKind = HttpBodyEmpty
    mText = ""
    mBytes = Array()
End Sub

Public Function Clone() As HttpBody
    Dim copy As New HttpBody

    Select Case mKind
    Case HttpBodyText
        copy.Text = mText
    Case HttpBodyBytes
        copy.SetBytes mBytes
    End Select
    Set Clone = copy
End Function

Private Sub Class_Initialize()
    Clear
End Sub
