VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpRetryPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMaxAttempts As Long
Private mBaseDelayMilliseconds As Long
Private mMaxDelayMilliseconds As Long
Private mJitterRatio As Double
Private mRespectRetryAfter As Boolean
Private mRetryNonIdempotentMethods As Boolean

Public Property Get MaxAttempts() As Long
    MaxAttempts = mMaxAttempts
End Property

Public Property Let MaxAttempts(ByVal value As Long)
    If value < 1 Or value > 100 Then HttpErrors.RaiseValidation "HttpRetryPolicy.MaxAttempts", "MaxAttempts must be between 1 and 100."
    mMaxAttempts = value
End Property

Public Property Get BaseDelayMilliseconds() As Long
    BaseDelayMilliseconds = mBaseDelayMilliseconds
End Property

Public Property Let BaseDelayMilliseconds(ByVal value As Long)
    If value < 0 Then HttpErrors.RaiseValidation "HttpRetryPolicy.BaseDelayMilliseconds", "Base delay cannot be negative."
    mBaseDelayMilliseconds = value
End Property

Public Property Get MaxDelayMilliseconds() As Long
    MaxDelayMilliseconds = mMaxDelayMilliseconds
End Property

Public Property Let MaxDelayMilliseconds(ByVal value As Long)
    If value < 0 Then HttpErrors.RaiseValidation "HttpRetryPolicy.MaxDelayMilliseconds", "Maximum delay cannot be negative."
    mMaxDelayMilliseconds = value
End Property

Public Property Get JitterRatio() As Double
    JitterRatio = mJitterRatio
End Property

Public Property Let JitterRatio(ByVal value As Double)
    If value < 0 Or value > 1 Then HttpErrors.RaiseValidation "HttpRetryPolicy.JitterRatio", "Jitter ratio must be between 0 and 1."
    mJitterRatio = value
End Property

Public Property Get RespectRetryAfter() As Boolean
    RespectRetryAfter = mRespectRetryAfter
End Property

Public Property Let RespectRetryAfter(ByVal value As Boolean)
    mRespectRetryAfter = value
End Property

Public Property Get RetryNonIdempotentMethods() As Boolean
    RetryNonIdempotentMethods = mRetryNonIdempotentMethods
End Property

Public Property Let RetryNonIdempotentMethods(ByVal value As Boolean)
    mRetryNonIdempotentMethods = value
End Property

Public Function Clone() As HttpRetryPolicy
    Dim copy As New HttpRetryPolicy

    Validate
    copy.MaxAttempts = mMaxAttempts
    copy.BaseDelayMilliseconds = mBaseDelayMilliseconds
    copy.MaxDelayMilliseconds = mMaxDelayMilliseconds
    copy.JitterRatio = mJitterRatio
    copy.RespectRetryAfter = mRespectRetryAfter
    copy.RetryNonIdempotentMethods = mRetryNonIdempotentMethods
    Set Clone = copy
End Function

Friend Sub Validate()
    If mMaxDelayMilliseconds < mBaseDelayMilliseconds Then
        HttpErrors.RaiseValidation "HttpRetryPolicy", "Maximum delay cannot be less than base delay."
    End If
End Sub

Friend Function AllowsMethod(ByVal Method As String) As Boolean
    Select Case UCase$(Trim$(Method))
    Case "GET", "HEAD", "PUT", "DELETE", "OPTIONS", "TRACE"
        AllowsMethod = True
    Case Else
        AllowsMethod = mRetryNonIdempotentMethods
    End Select
End Function

Friend Function AllowsStatus(ByVal statusCode As Long) As Boolean
    Select Case statusCode
    Case 408, 429, 500, 502, 503, 504
        AllowsStatus = True
    End Select
End Function

Friend Function AllowsError(ByVal category As HttpErrorCategory) As Boolean
    Select Case category
    Case HttpErrorDns, HttpErrorConnection, HttpErrorTimeout, HttpErrorIo
        AllowsError = True
    End Select
End Function

Friend Function BackoffMilliseconds(ByVal completedAttempt As Long, ByVal jitterUnit As Double) As Long
    Dim delay As Double
    Dim index As Long

    Validate
    If completedAttempt < 1 Then HttpErrors.RaiseValidation "HttpRetryPolicy.BackoffMilliseconds", "Completed attempt must be positive."
    If jitterUnit < 0 Or jitterUnit > 1 Then HttpErrors.RaiseValidation "HttpRetryPolicy.BackoffMilliseconds", "Jitter sample must be between 0 and 1."

    delay = CDbl(mBaseDelayMilliseconds)
    For index = 2 To completedAttempt
        If delay >= CDbl(mMaxDelayMilliseconds) / 2# Then
            delay = CDbl(mMaxDelayMilliseconds)
            Exit For
        End If
        delay = delay * 2#
    Next index
    delay = delay * (1# - mJitterRatio + 2# * mJitterRatio * jitterUnit)
    If delay < 0 Then delay = 0
    If delay > mMaxDelayMilliseconds Then delay = mMaxDelayMilliseconds
    BackoffMilliseconds = CLng(Fix(delay + 0.5))
End Function

Private Sub Class_Initialize()
    mMaxAttempts = 3
    mBaseDelayMilliseconds = 200
    mMaxDelayMilliseconds = 30000
    mJitterRatio = 0.2
    mRespectRetryAfter = True
End Sub
