VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordingProgressSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IHttpProgressSink

Private mCallCount As Long
Private mLastBytes As Currency
Private mLastTotal As Currency
Private mCancelAfter As Currency
Private mCancellationToken As HttpCancellationToken

Public Property Get CallCount() As Long
    CallCount = mCallCount
End Property

Public Property Get LastBytesTransferred() As Currency
    LastBytesTransferred = mLastBytes
End Property

Public Property Get LastTotalBytes() As Currency
    LastTotalBytes = mLastTotal
End Property

Public Property Let CancelAfter(ByVal value As Currency)
    If value < -1 Then HttpErrors.RaiseValidation "RecordingProgressSink.CancelAfter", "CancelAfter must be -1 or greater."
    mCancelAfter = value
End Property

Public Property Set CancellationToken(ByVal value As HttpCancellationToken)
    Set mCancellationToken = value
End Property

Private Sub IHttpProgressSink_OnProgress(ByVal BytesTransferred As Currency, ByVal TotalBytes As Currency)
    mCallCount = mCallCount + 1
    mLastBytes = BytesTransferred
    mLastTotal = TotalBytes
    If mCancelAfter >= 0 And BytesTransferred >= mCancelAfter Then
        If Not mCancellationToken Is Nothing Then mCancellationToken.Cancel
    End If
End Sub

Private Sub Class_Initialize()
    mCancelAfter = -1
    mLastTotal = -1
End Sub
