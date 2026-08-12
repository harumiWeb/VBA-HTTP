VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Execute(ByVal Request As HttpRequest) As HttpResponse
    Err.Raise 5, "IHttpTransport.Execute", "Interface method must be implemented."
End Function
