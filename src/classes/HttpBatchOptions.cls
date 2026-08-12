VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpBatchOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMaxConcurrency As Long
Private mPollIntervalMilliseconds As Long
Private mYieldToHost As Boolean
Private mYieldIntervalMilliseconds As Long
Private mRequestDeadlineMilliseconds As Long
Private mCancellationToken As HttpCancellationToken
Private mRetryPolicy As HttpRetryPolicy
Private mTotalDeadlineMilliseconds As Long

Public Property Get MaxConcurrency() As Long
    MaxConcurrency = mMaxConcurrency
End Property

Public Property Let MaxConcurrency(ByVal value As Long)
    If value < 1 Or value > 64 Then HttpErrors.RaiseValidation "HttpBatchOptions.MaxConcurrency", "MaxConcurrency must be between 1 and 64."
    mMaxConcurrency = value
End Property

Public Property Get PollIntervalMilliseconds() As Long
    PollIntervalMilliseconds = mPollIntervalMilliseconds
End Property

Public Property Let PollIntervalMilliseconds(ByVal value As Long)
    If value < 1 Or value > 1000 Then HttpErrors.RaiseValidation "HttpBatchOptions.PollIntervalMilliseconds", "Poll interval must be between 1 and 1000 milliseconds."
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
    If value < 1 Or value > 1000 Then HttpErrors.RaiseValidation "HttpBatchOptions.YieldIntervalMilliseconds", "Yield interval must be between 1 and 1000 milliseconds."
    mYieldIntervalMilliseconds = value
End Property

Public Property Get RequestDeadlineMilliseconds() As Long
    RequestDeadlineMilliseconds = mRequestDeadlineMilliseconds
End Property

Public Property Let RequestDeadlineMilliseconds(ByVal value As Long)
    If value < 0 Then HttpErrors.RaiseValidation "HttpBatchOptions.RequestDeadlineMilliseconds", "Request deadline cannot be negative."
    mRequestDeadlineMilliseconds = value
End Property

Public Property Get CancellationToken() As HttpCancellationToken
    Set CancellationToken = mCancellationToken
End Property

Public Property Set CancellationToken(ByVal value As HttpCancellationToken)
    Set mCancellationToken = value
End Property

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
    If value < 0 Then HttpErrors.RaiseValidation "HttpBatchOptions.TotalDeadlineMilliseconds", "Total deadline cannot be negative."
    mTotalDeadlineMilliseconds = value
End Property

Public Function Clone() As HttpBatchOptions
    Dim copy As New HttpBatchOptions

    copy.MaxConcurrency = mMaxConcurrency
    copy.PollIntervalMilliseconds = mPollIntervalMilliseconds
    copy.YieldToHost = mYieldToHost
    copy.YieldIntervalMilliseconds = mYieldIntervalMilliseconds
    copy.RequestDeadlineMilliseconds = mRequestDeadlineMilliseconds
    Set copy.CancellationToken = mCancellationToken
    If Not mRetryPolicy Is Nothing Then Set copy.RetryPolicy = mRetryPolicy
    copy.TotalDeadlineMilliseconds = mTotalDeadlineMilliseconds
    Set Clone = copy
End Function

Private Sub Class_Initialize()
    mMaxConcurrency = 8
    mPollIntervalMilliseconds = 5
    mYieldToHost = True
    mYieldIntervalMilliseconds = 20
End Sub
