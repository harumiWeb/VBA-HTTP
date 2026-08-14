VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpMultipartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mBoundary As String
Private mParts As New Collection

Public Property Get Boundary() As String
    Boundary = mBoundary
End Property

Public Property Let Boundary(ByVal value As String)
    ValidateBoundary value
    mBoundary = value
End Property

Public Property Get Count() As Long
    Count = mParts.Count
End Property

Public Sub AddField(ByVal Name As String, ByVal Value As String)
    Dim part As New HttpMultipartPart

    ValidateHeaderToken Name, "HttpMultipartForm.AddField"
    part.InitializeField Name, Value
    mParts.Add part
End Sub

Public Sub AddFile(ByVal Name As String, ByVal FilePath As String, Optional ByVal FileName As String = "", Optional ByVal ContentType As String = "application/octet-stream")
    Dim part As New HttpMultipartPart
    Dim normalizedPath As String
    Dim fileSystem As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem Is Nothing Then
        Err.Raise vbObjectError + 736, "HttpMultipartForm.AddFile", "File system service is unavailable."
        Exit Sub
    End If
    ValidateHeaderToken Name, "HttpMultipartForm.AddFile"
    ValidateHeaderValue ContentType, "HttpMultipartForm.AddFile"
    normalizedPath = WinHttpUploadFileSystem.NormalizeSource(FilePath)
    If Len(FileName) = 0 Then FileName = fileSystem.GetFileName(normalizedPath)
    ValidateHeaderToken FileName, "HttpMultipartForm.AddFile"
    If Len(ContentType) = 0 Then HttpErrors.RaiseValidation "HttpMultipartForm.AddFile", "Content type cannot be empty."
    part.InitializeFile Name, normalizedPath, FileName, ContentType
    mParts.Add part
End Sub

Public Function PartAt(ByVal index As Long) As HttpMultipartPart
    If index < 1 Or index > mParts.Count Then HttpErrors.RaiseValidation "HttpMultipartForm.PartAt", "Part index is out of range."
    Set PartAt = mParts(index)
End Function

Friend Function Clone() As HttpMultipartForm
    Dim copy As New HttpMultipartForm
    Dim part As HttpMultipartPart
    Dim clonedPart As HttpMultipartPart
    Dim index As Long

    copy.Boundary = mBoundary
    For index = 1 To mParts.Count
        Set part = mParts(index)
        Set clonedPart = part.Clone()
        copy.AddPart clonedPart
    Next index
    Set Clone = copy
End Function

Friend Sub AddPart(ByVal part As HttpMultipartPart)
    If part Is Nothing Then HttpErrors.RaiseValidation "HttpMultipartForm.AddPart", "Multipart part cannot be Nothing."
    mParts.Add part
End Sub

Private Sub ValidateBoundary(ByVal value As String)
    Dim index As Long
    Dim code As Long
    Dim character As String

    If Len(value) = 0 Or Len(value) > 70 Then HttpErrors.RaiseValidation "HttpMultipartForm.Boundary", "Boundary must contain between 1 and 70 characters."
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        code = AscW(character)
        If code < 0 Then code = code + 65536
        If Not ((code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or character = "-") Then
            HttpErrors.RaiseValidation "HttpMultipartForm.Boundary", "Boundary contains an unsafe character."
        End If
    Next index
End Sub

Private Sub ValidateHeaderToken(ByVal value As String, ByVal Source As String)
    Dim index As Long
    Dim code As Long
    Dim character As String

    If Len(value) = 0 Then HttpErrors.RaiseValidation Source, "Multipart header token cannot be empty."
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        code = AscW(character)
        If code < 0 Then code = code + 65536
        If code < 33 Or code > 126 Or character = """" Or character = "\" Then
            HttpErrors.RaiseValidation Source, "Multipart header token contains an unsafe character."
        End If
    Next index
End Sub

Private Sub ValidateHeaderValue(ByVal value As String, ByVal Source As String)
    Dim index As Long
    Dim code As Long

    For index = 1 To Len(value)
        code = AscW(Mid$(value, index, 1))
        If code < 0 Then code = code + 65536
        If code = 13 Or code = 10 Or code < 32 Or code = 127 Then HttpErrors.RaiseValidation Source, "Multipart header value contains a control character."
    Next index
End Sub

Private Sub Class_Initialize()
    Set mParts = New Collection
    mBoundary = WinHttpUploadEncoding.NextBoundary()
End Sub
