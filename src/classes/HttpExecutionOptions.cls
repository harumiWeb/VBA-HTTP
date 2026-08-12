VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpExecutionOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mRetryPolicy As HttpRetryPolicy
Private mTotalDeadlineMilliseconds As Long
Private mCancellationToken As HttpCancellationToken
Private mPollIntervalMilliseconds As Long
Private mYieldToHost As Boolean
Private mYieldIntervalMilliseconds As Long

Public Property Get RetryPolicy() As HttpRetryPolicy
    Set RetryPolicy = mRetryPolicy
End Property

Public Property Set RetryPolicy(ByVal value As HttpRetryPolicy)
    If value Is Nothing Then
        Set mRetryPolicy = Nothing
    Else
        Set mRetryPolicy = value.Clone()
    End If
End Property

Public Property Get TotalDeadlineMilliseconds() As Long
    TotalDeadlineMilliseconds = mTotalDeadlineMilliseconds
End Property

Public Property Let TotalDeadlineMilliseconds(ByVal value As Long)
    If value < 0 Then HttpErrors.RaiseValidation "HttpExecutionOptions.TotalDeadlineMilliseconds", "Total deadline cannot be negative."
    mTotalDeadlineMilliseconds = value
End Property

Public Property Get CancellationToken() As HttpCancellationToken
    Set CancellationToken = mCancellationToken
End Property

Public Property Set CancellationToken(ByVal value As HttpCancellationToken)
    Set mCancellationToken = value
End Property

Public Property Get PollIntervalMilliseconds() As Long
    PollIntervalMilliseconds = mPollIntervalMilliseconds
End Property

Public Property Let PollIntervalMilliseconds(ByVal value As Long)
    If value < 1 Or value > 1000 Then HttpErrors.RaiseValidation "HttpExecutionOptions.PollIntervalMilliseconds", "Poll interval must be between 1 and 1000 milliseconds."
    mPollIntervalMilliseconds = value
End Property

Public Property Get YieldToHost() As Boolean
    YieldToHost = mYieldToHost
End Property

Public Property Let YieldToHost(ByVal value As Boolean)
    mYieldToHost = value
End Property

Public Property Get YieldIntervalMilliseconds() As Long
    YieldIntervalMilliseconds = mYieldIntervalMilliseconds
End Property

Public Property Let YieldIntervalMilliseconds(ByVal value As Long)
    If value < 1 Or value > 1000 Then HttpErrors.RaiseValidation "HttpExecutionOptions.YieldIntervalMilliseconds", "Yield interval must be between 1 and 1000 milliseconds."
    mYieldIntervalMilliseconds = value
End Property

Public Function Clone() As HttpExecutionOptions
    Dim copy As New HttpExecutionOptions

    If Not mRetryPolicy Is Nothing Then Set copy.RetryPolicy = mRetryPolicy
    copy.TotalDeadlineMilliseconds = mTotalDeadlineMilliseconds
    Set copy.CancellationToken = mCancellationToken
    copy.PollIntervalMilliseconds = mPollIntervalMilliseconds
    copy.YieldToHost = mYieldToHost
    copy.YieldIntervalMilliseconds = mYieldIntervalMilliseconds
    Set Clone = copy
End Function

Private Sub Class_Initialize()
    mPollIntervalMilliseconds = 5
    mYieldToHost = True
    mYieldIntervalMilliseconds = 20
End Sub
