VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeCookieClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IHttpCookieClock

Private mNowUtc As Date

Public Property Let NowUtc(ByVal value As Date)
    mNowUtc = value
End Property

Private Function IHttpCookieClock_UtcNow() As Date
    IHttpCookieClock_UtcNow = mNowUtc
End Function

Private Sub Class_Initialize()
    mNowUtc = DateSerial(2026, 1, 1)
End Sub
