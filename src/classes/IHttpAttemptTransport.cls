VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpAttemptTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExecuteAttempt(ByVal Request As HttpRequest, ByVal Options As HttpAttemptOptions) As HttpResponse
    Err.Raise 5, "IHttpAttemptTransport.ExecuteAttempt", "Interface method must be implemented."
End Function
