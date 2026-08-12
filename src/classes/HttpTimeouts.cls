VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpTimeouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mResolveMilliseconds As Long
Private mConnectMilliseconds As Long
Private mSendMilliseconds As Long
Private mReceiveMilliseconds As Long

Public Property Get ResolveMilliseconds() As Long
    ResolveMilliseconds = mResolveMilliseconds
End Property

Public Property Let ResolveMilliseconds(ByVal value As Long)
    ValidateTimeout value
    mResolveMilliseconds = value
End Property

Public Property Get ConnectMilliseconds() As Long
    ConnectMilliseconds = mConnectMilliseconds
End Property

Public Property Let ConnectMilliseconds(ByVal value As Long)
    ValidateTimeout value
    mConnectMilliseconds = value
End Property

Public Property Get SendMilliseconds() As Long
    SendMilliseconds = mSendMilliseconds
End Property

Public Property Let SendMilliseconds(ByVal value As Long)
    ValidateTimeout value
    mSendMilliseconds = value
End Property

Public Property Get ReceiveMilliseconds() As Long
    ReceiveMilliseconds = mReceiveMilliseconds
End Property

Public Property Let ReceiveMilliseconds(ByVal value As Long)
    ValidateTimeout value
    mReceiveMilliseconds = value
End Property

Public Function Clone() As HttpTimeouts
    Dim copy As New HttpTimeouts

    copy.ResolveMilliseconds = mResolveMilliseconds
    copy.ConnectMilliseconds = mConnectMilliseconds
    copy.SendMilliseconds = mSendMilliseconds
    copy.ReceiveMilliseconds = mReceiveMilliseconds
    Set Clone = copy
End Function

Private Sub Class_Initialize()
    mResolveMilliseconds = 5000
    mConnectMilliseconds = 5000
    mSendMilliseconds = 30000
    mReceiveMilliseconds = 300000
End Sub

Private Sub ValidateTimeout(ByVal value As Long)
    If value < 0 Then
        HttpErrors.RaiseValidation "HttpTimeouts", "Timeout values must be non-negative."
    End If
End Sub
