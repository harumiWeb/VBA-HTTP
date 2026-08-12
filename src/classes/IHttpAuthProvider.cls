VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpAuthProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Apply(ByVal Request As HttpRequest)
    Err.Raise 5, "IHttpAuthProvider.Apply", "Interface method must be implemented."
End Sub

Public Function Clone() As IHttpAuthProvider
    Err.Raise 5, "IHttpAuthProvider.Clone", "Interface method must be implemented."
End Function
