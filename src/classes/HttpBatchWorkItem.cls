VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpBatchWorkItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIndex As Long
Private mRequest As HttpRequest
Private mAttempt As Long
Private mNextEligibleMilliseconds As Double
Private mRunning As Boolean
Private mResult As HttpBatchItem

Friend Property Get Index() As Long
    Index = mIndex
End Property

Friend Property Get RequestSnapshot() As HttpRequest
    Set RequestSnapshot = mRequest
End Property

Friend Property Get Attempt() As Long
    Attempt = mAttempt
End Property

Friend Property Get IsRunning() As Boolean
    IsRunning = mRunning
End Property

Friend Property Get IsComplete() As Boolean
    IsComplete = Not mResult Is Nothing
End Property

Friend Property Get TerminalItem() As HttpBatchItem
    Set TerminalItem = mResult
End Property

Friend Sub Initialize(ByVal itemIndex As Long, ByVal value As HttpRequest)
    If itemIndex < 1 Then HttpErrors.RaiseValidation "HttpBatchWorkItem.Initialize", "Batch item index must be positive."
    If value Is Nothing Then HttpErrors.RaiseValidation "HttpBatchWorkItem.Initialize", "Batch request cannot be Nothing."
    mIndex = itemIndex
    Set mRequest = value
End Sub

Friend Function IsEligible(ByVal elapsedMilliseconds As Double) As Boolean
    IsEligible = (Not IsComplete And Not mRunning And elapsedMilliseconds >= mNextEligibleMilliseconds)
End Function

Friend Sub BeginAttempt()
    If IsComplete Or mRunning Then Err.Raise HttpErrProtocol, "HttpBatchWorkItem.BeginAttempt", "Batch work item is not pending."
    mAttempt = mAttempt + 1
    mRunning = True
End Sub

Friend Sub ScheduleRetry(ByVal elapsedMilliseconds As Double, ByVal delayMilliseconds As Long)
    If delayMilliseconds < 0 Then HttpErrors.RaiseValidation "HttpBatchWorkItem.ScheduleRetry", "Retry delay cannot be negative."
    mRunning = False
    mNextEligibleMilliseconds = elapsedMilliseconds + CDbl(delayMilliseconds)
End Sub

Friend Sub CompleteSuccess(ByVal response As HttpResponse)
    Dim item As New HttpBatchItem

    item.InitializeSuccess mIndex, response
    Set mResult = item
    mRunning = False
End Sub

Friend Sub CompleteFailure(ByVal errorNumber As Long, ByVal errorSource As String, ByVal errorDescription As String)
    Dim item As New HttpBatchItem

    item.InitializeFailure mIndex, errorNumber, errorSource, errorDescription
    Set mResult = item
    mRunning = False
End Sub

Friend Sub CompleteCancelled()
    Dim item As New HttpBatchItem

    item.InitializeCancelled mIndex
    Set mResult = item
    mRunning = False
End Sub

Friend Sub MarkNotRunning()
    mRunning = False
End Sub
