VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpUploadSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ContentLength() As Currency
    Err.Raise 5, "IHttpUploadSource.ContentLength", "Interface method must be implemented."
End Function

Public Function ContentType() As String
    Err.Raise 5, "IHttpUploadSource.ContentType", "Interface method must be implemented."
End Function

Public Sub WriteTo(ByVal Writer As WinHttpUploadStreamWriter)
    Err.Raise 5, "IHttpUploadSource.WriteTo", "Interface method must be implemented."
End Sub
