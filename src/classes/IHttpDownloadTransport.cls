VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpDownloadTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DownloadFile(ByVal Request As HttpRequest, ByVal DestinationPath As String, ByVal Options As HttpExecutionOptions, Optional ByVal Progress As IHttpProgressSink = Nothing) As HttpDownloadResult
    Err.Raise 5, "IHttpDownloadTransport.DownloadFile", "Interface method must be implemented."
End Function
