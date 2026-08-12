VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpMultipartUploadSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IHttpUploadSource

Private mEncoder As WinHttpMultipartEncoder
Private mLength As Currency
Private mContentType As String

Friend Sub Initialize(ByVal Form As HttpMultipartForm)
    Set mEncoder = New WinHttpMultipartEncoder
    mEncoder.Initialize Form
    mLength = mEncoder.ContentLength
    mContentType = mEncoder.ContentType
End Sub

Private Function IHttpUploadSource_ContentLength() As Currency
    IHttpUploadSource_ContentLength = mLength
End Function

Private Function IHttpUploadSource_ContentType() As String
    IHttpUploadSource_ContentType = mContentType
End Function

Private Sub IHttpUploadSource_WriteTo(ByVal Writer As WinHttpUploadStreamWriter)
    mEncoder.WriteTo Writer
End Sub
