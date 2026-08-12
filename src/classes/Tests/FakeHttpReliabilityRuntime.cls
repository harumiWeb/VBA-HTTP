VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeHttpReliabilityRuntime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IHttpReliabilityRuntime

Private mNowMilliseconds As Currency
Private mUtcNow As Date
Private mJitterUnit As Double
Private mPausedMilliseconds As Long
Private mYieldCount As Long
Private mCancellationToken As HttpCancellationToken
Private mCancelAtMilliseconds As Long

Public Property Get PausedMilliseconds() As Long
    PausedMilliseconds = mPausedMilliseconds
End Property

Public Property Get YieldCount() As Long
    YieldCount = mYieldCount
End Property

Public Property Let UtcNow(ByVal value As Date)
    mUtcNow = value
End Property

Public Property Let JitterUnit(ByVal value As Double)
    If value < 0 Or value > 1 Then Err.Raise 5, "FakeHttpReliabilityRuntime.JitterUnit", "Jitter must be between 0 and 1."
    mJitterUnit = value
End Property

Public Sub CancelAfter(ByVal milliseconds As Long, ByVal token As HttpCancellationToken)
    mCancelAtMilliseconds = milliseconds
    Set mCancellationToken = token
End Sub

Private Function IHttpReliabilityRuntime_CounterValue() As Currency
    IHttpReliabilityRuntime_CounterValue = mNowMilliseconds
End Function

Private Function IHttpReliabilityRuntime_ElapsedMilliseconds(ByVal started As Currency, ByVal finished As Currency) As Double
    IHttpReliabilityRuntime_ElapsedMilliseconds = CDbl(finished - started)
End Function

Private Function IHttpReliabilityRuntime_UtcNow() As Date
    IHttpReliabilityRuntime_UtcNow = mUtcNow
End Function

Private Function IHttpReliabilityRuntime_NextJitterUnit() As Double
    IHttpReliabilityRuntime_NextJitterUnit = mJitterUnit
End Function

Private Sub IHttpReliabilityRuntime_Pause(ByVal milliseconds As Long)
    mNowMilliseconds = mNowMilliseconds + milliseconds
    mPausedMilliseconds = mPausedMilliseconds + milliseconds
    If Not mCancellationToken Is Nothing And mCancelAtMilliseconds > 0 Then
        If mPausedMilliseconds >= mCancelAtMilliseconds Then mCancellationToken.Cancel
    End If
End Sub

Private Sub IHttpReliabilityRuntime_YieldHost()
    mYieldCount = mYieldCount + 1
End Sub

Private Sub Class_Initialize()
    mUtcNow = DateSerial(2026, 1, 1)
    mJitterUnit = 0.5
End Sub
