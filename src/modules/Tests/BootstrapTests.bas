Attribute VB_Name = "BootstrapTests"
Option Explicit

'@Tag("smoke")
Public Sub Test_XlflowHarness_ExecutesSmokeTest()
    XlflowAssert.AssertTrue True, "xlflow should execute workbook tests"
End Sub

Public Sub Test_DebugLog_AcceptsEmptyArgumentList()
    XlflowDebug.Log
End Sub

Public Sub Test_DebugLog_AcceptsValues()
    XlflowDebug.Log "debug", 1, True
End Sub
