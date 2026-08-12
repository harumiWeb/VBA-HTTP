VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReentrantBatchTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHttpTransport
Implements IHttpBatchTransport

Private mClient As HttpClient

Public Property Set Client(ByVal value As HttpClient)
    Set mClient = value
End Property

Private Function IHttpTransport_Execute(ByVal Request As HttpRequest) As HttpResponse
    Err.Raise HttpErrProtocol, "ReentrantBatchTransport.Execute", "Synchronous execution is not used by this test transport."
End Function

Private Function IHttpBatchTransport_ExecuteMany(ByVal Requests As Collection, ByVal Options As HttpBatchOptions) As HttpBatchResult
    Dim Request As New HttpRequest

    Request.Url = "http://127.0.0.1:1/status/204"
    Call mClient.Execute(Request)
End Function
