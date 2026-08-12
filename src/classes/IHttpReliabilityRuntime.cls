VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpReliabilityRuntime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CounterValue() As Currency
    Err.Raise 5, "IHttpReliabilityRuntime.CounterValue", "Interface method must be implemented."
End Function

Public Function ElapsedMilliseconds(ByVal started As Currency, ByVal finished As Currency) As Double
    Err.Raise 5, "IHttpReliabilityRuntime.ElapsedMilliseconds", "Interface method must be implemented."
End Function

Public Function UtcNow() As Date
    Err.Raise 5, "IHttpReliabilityRuntime.UtcNow", "Interface method must be implemented."
End Function

Public Function NextJitterUnit() As Double
    Err.Raise 5, "IHttpReliabilityRuntime.NextJitterUnit", "Interface method must be implemented."
End Function

Public Sub Pause(ByVal milliseconds As Long)
    Err.Raise 5, "IHttpReliabilityRuntime.Pause", "Interface method must be implemented."
End Sub

Public Sub YieldHost()
    Err.Raise 5, "IHttpReliabilityRuntime.YieldHost", "Interface method must be implemented."
End Sub
