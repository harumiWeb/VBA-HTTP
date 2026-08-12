VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpComOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIndex As Long
Private mBackend As Object
Private mStarted As Currency
Private mDeadlineMilliseconds As Long

Friend Property Get Index() As Long
    Index = mIndex
End Property

Friend Property Get Backend() As Object
    Set Backend = mBackend
End Property

Friend Property Get Started() As Currency
    Started = mStarted
End Property

Friend Sub Initialize(ByVal itemIndex As Long, ByVal backend As Object, ByVal started As Currency, ByVal deadlineMilliseconds As Long)
    mIndex = itemIndex
    Set mBackend = backend
    mStarted = started
    mDeadlineMilliseconds = deadlineMilliseconds
End Sub

Friend Function DeadlineExceeded(ByVal now As Currency) As Boolean
    If mDeadlineMilliseconds = 0 Then Exit Function
    DeadlineExceeded = (HttpTiming.ElapsedMilliseconds(mStarted, now) >= mDeadlineMilliseconds)
End Function

Friend Sub ReleaseBackend()
    Set mBackend = Nothing
End Sub
