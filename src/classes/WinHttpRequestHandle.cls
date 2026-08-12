VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpRequestHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mParent As WinHttpConnectionHandle
#If VBA7 Then
Private mHandle As LongPtr

Friend Sub Initialize(ByVal HandleValue As LongPtr, ByVal Parent As WinHttpConnectionHandle)
    If HandleValue = 0 Then Err.Raise HttpErrIo, "WinHttpRequestHandle.Initialize", "WinHTTP request handle was not created."
    If Parent Is Nothing Then Err.Raise HttpErrValidation, "WinHttpRequestHandle.Initialize", "Request parent cannot be Nothing."
    Set mParent = Parent
    mHandle = HandleValue
End Sub

Friend Property Get Value() As LongPtr
    Value = mHandle
End Property
#Else
Private mHandle As Long

Friend Sub Initialize(ByVal HandleValue As Long, ByVal Parent As WinHttpConnectionHandle)
    If HandleValue = 0 Then Err.Raise HttpErrIo, "WinHttpRequestHandle.Initialize", "WinHTTP request handle was not created."
    If Parent Is Nothing Then Err.Raise HttpErrValidation, "WinHttpRequestHandle.Initialize", "Request parent cannot be Nothing."
    Set mParent = Parent
    mHandle = HandleValue
End Sub

Friend Property Get Value() As Long
    Value = mHandle
End Property
#End If

Friend Property Get IsOpen() As Boolean
    IsOpen = (mHandle <> 0)
End Property

Public Sub ReleaseHandle()
    If mHandle = 0 Then Exit Sub
    Call WinHttpNativeApi.CloseHandle(mHandle)
    mHandle = 0
    Set mParent = Nothing
End Sub

Private Sub Class_Terminate()
    ReleaseHandle
End Sub
