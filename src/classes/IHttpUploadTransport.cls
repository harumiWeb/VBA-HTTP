VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpUploadTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function UploadFile(ByVal Request As HttpRequest, ByVal SourcePath As String, ByVal ContentType As String, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing) As HttpUploadResult
    Err.Raise 5, "IHttpUploadTransport.UploadFile", "Interface method must be implemented."
End Function

Public Function UploadMultipart(ByVal Request As HttpRequest, ByVal Form As HttpMultipartForm, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing) As HttpUploadResult
    Err.Raise 5, "IHttpUploadTransport.UploadMultipart", "Interface method must be implemented."
End Function
