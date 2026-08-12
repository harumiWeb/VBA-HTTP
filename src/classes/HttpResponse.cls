VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mStatusCode As Long
Private mReasonPhrase As String
Private mHeaders As New HttpHeaders
Private mBody As New HttpBody
Private mElapsedMilliseconds As Double
Private mTextMaterialized As Boolean
Private mText As String

Public Property Get StatusCode() As Long
    StatusCode = mStatusCode
End Property

Public Property Get ReasonPhrase() As String
    ReasonPhrase = mReasonPhrase
End Property

Public Property Get Headers() As HttpHeaders
    Set Headers = mHeaders
End Property

Public Property Get Body() As HttpBody
    Set Body = mBody
End Property

Public Property Get ElapsedMilliseconds() As Double
    ElapsedMilliseconds = mElapsedMilliseconds
End Property

Public Property Get IsSuccess() As Boolean
    IsSuccess = (mStatusCode >= 200 And mStatusCode <= 299)
End Property

Public Property Get Text() As String
    If Not mTextMaterialized Then
        mText = DecodeBodyText()
        mTextMaterialized = True
    End If
    Text = mText
End Property

Public Sub Initialize(ByVal statusCode As Long, Optional ByVal reasonPhrase As String = "", Optional ByVal Headers As HttpHeaders = Nothing, Optional ByVal Body As HttpBody = Nothing, Optional ByVal elapsedMilliseconds As Double = 0)
    If statusCode < 100 Or statusCode > 599 Then
        HttpErrors.RaiseValidation "HttpResponse.Initialize", "HTTP status code must be between 100 and 599."
    End If
    If elapsedMilliseconds < 0 Then
        HttpErrors.RaiseValidation "HttpResponse.Initialize", "Elapsed milliseconds cannot be negative."
    End If

    mStatusCode = statusCode
    mReasonPhrase = reasonPhrase
    If Headers Is Nothing Then
        Set mHeaders = New HttpHeaders
    Else
        Set mHeaders = Headers.Clone()
    End If
    If Body Is Nothing Then
        Set mBody = New HttpBody
    Else
        Set mBody = Body.Clone()
    End If
    mElapsedMilliseconds = elapsedMilliseconds
    mTextMaterialized = False
    mText = ""
End Sub

Public Sub RaiseForStatus()
    If Not IsSuccess Then
        HttpErrors.RaiseStatus "HttpResponse.RaiseForStatus", mStatusCode
    End If
End Sub

Public Function Clone() As HttpResponse
    Dim copy As New HttpResponse

    copy.Initialize mStatusCode, mReasonPhrase, mHeaders, mBody, mElapsedMilliseconds
    Set Clone = copy
End Function

Private Function DecodeBodyText() As String
    Dim charset As String
    Dim bytes As Variant

    If mBody.Kind = HttpBodyText Then
        DecodeBodyText = mBody.Text
        Exit Function
    End If
    If mBody.IsEmpty Then Exit Function

    charset = ResponseCharset()
    bytes = mBody.Bytes
    Select Case charset
    Case "", "utf-8", "utf8"
        DecodeBodyText = HttpEncoding.DecodeUtf8(bytes)
    Case "us-ascii", "ascii"
        DecodeBodyText = HttpEncoding.DecodeAscii(bytes)
    Case Else
        HttpErrors.RaiseValidation "HttpResponse.Text", "Unsupported response charset '" & charset & "'."
    End Select
End Function

Private Function ResponseCharset() As String
    Dim contentType As String
    Dim parts As Variant
    Dim part As Variant
    Dim candidate As String

    contentType = mHeaders.GetValue("Content-Type")
    If Len(contentType) = 0 Then Exit Function
    parts = Split(contentType, ";")
    For Each part In parts
        candidate = Trim$(CStr(part))
        If LCase$(Left$(candidate, 8)) = "charset=" Then
            candidate = Trim$(Mid$(candidate, 9))
            If Len(candidate) >= 2 And Left$(candidate, 1) = """" And Right$(candidate, 1) = """" Then
                candidate = Mid$(candidate, 2, Len(candidate) - 2)
            End If
            ResponseCharset = LCase$(candidate)
            Exit Function
        End If
    Next part
End Function
