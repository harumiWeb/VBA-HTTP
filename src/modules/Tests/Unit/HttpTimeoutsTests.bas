Attribute VB_Name = "HttpTimeoutsTests"
Option Explicit

Public Sub Test_Timeouts_HasIndependentDefaults()
    Dim first As New HttpTimeouts
    Dim second As New HttpTimeouts

    first.ConnectMilliseconds = 42

    XlflowAssert.AssertEquals 5000, first.ResolveMilliseconds
    XlflowAssert.AssertEquals 42, first.ConnectMilliseconds
    XlflowAssert.AssertEquals 30000, first.SendMilliseconds
    XlflowAssert.AssertEquals 300000, first.ReceiveMilliseconds
    XlflowAssert.AssertEquals 5000, second.ConnectMilliseconds
End Sub

Public Sub Test_Timeouts_CloneIsIndependent()
    Dim original As New HttpTimeouts
    Dim copy As HttpTimeouts

    Set copy = original.Clone()
    copy.ReceiveMilliseconds = 99

    XlflowAssert.AssertEquals 300000, original.ReceiveMilliseconds
    XlflowAssert.AssertEquals 99, copy.ReceiveMilliseconds
End Sub

'@ExpectedError(-2147200503, "Timeout values must be non-negative.", "HttpTimeouts")
Public Sub Test_Timeouts_RejectsNegativeValue()
    Dim timeouts As New HttpTimeouts
    timeouts.SendMilliseconds = -1
End Sub
