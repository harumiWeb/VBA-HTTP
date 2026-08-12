VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpProgressSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub OnProgress(ByVal BytesTransferred As Currency, ByVal TotalBytes As Currency)
    Err.Raise 5, "IHttpProgressSink.OnProgress", "Interface method must be implemented."
End Sub
