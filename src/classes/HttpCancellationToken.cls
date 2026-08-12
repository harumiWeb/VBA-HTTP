VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpCancellationToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCancellationRequested As Boolean

Public Property Get IsCancellationRequested() As Boolean
    IsCancellationRequested = mCancellationRequested
End Property

Public Sub Cancel()
    mCancellationRequested = True
End Sub
