VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpSystemReliabilityRuntime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IHttpReliabilityRuntime

Private Function IHttpReliabilityRuntime_CounterValue() As Currency
    IHttpReliabilityRuntime_CounterValue = HttpTiming.CounterValue()
End Function

Private Function IHttpReliabilityRuntime_ElapsedMilliseconds(ByVal started As Currency, ByVal finished As Currency) As Double
    IHttpReliabilityRuntime_ElapsedMilliseconds = HttpTiming.ElapsedMilliseconds(started, finished)
End Function

Private Function IHttpReliabilityRuntime_UtcNow() As Date
    IHttpReliabilityRuntime_UtcNow = HttpTiming.UtcNow()
End Function

Private Function IHttpReliabilityRuntime_NextJitterUnit() As Double
    Dim sample As Double

    sample = Abs(CDbl(HttpTiming.CounterValue()))
    IHttpReliabilityRuntime_NextJitterUnit = sample - Fix(sample)
End Function

Private Sub IHttpReliabilityRuntime_Pause(ByVal milliseconds As Long)
    HttpTiming.Pause milliseconds
End Sub

Private Sub IHttpReliabilityRuntime_YieldHost()
    ' Reliability waits share ADR-0005's explicit cooperative host boundary.
    DoEvents ' xlflow:disable-line VB007
End Sub
