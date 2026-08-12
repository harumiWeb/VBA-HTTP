Attribute VB_Name = "WinHttpNativeTests"
Option Explicit

Public Sub Test_NativeErrorMapping_ClassifiesConnectionFailures()
    XlflowAssert.AssertEquals HttpErrorConnection, WinHttpErrorMapping.CategoryFromCode(12029)
    XlflowAssert.AssertEquals HttpErrorConnection, WinHttpErrorMapping.CategoryFromCode(12030)
    XlflowAssert.AssertEquals HttpErrorConnection, WinHttpErrorMapping.CategoryFromCode(12031)
End Sub

Public Sub Test_NativeErrorMapping_ClassifiesTlsFailures()
    XlflowAssert.AssertEquals HttpErrorTls, WinHttpErrorMapping.CategoryFromCode(12037)
    XlflowAssert.AssertEquals HttpErrorTls, WinHttpErrorMapping.CategoryFromCode(12175)
End Sub

Public Sub Test_NativeErrorMapping_ClassifiesProtocolFailures()
    XlflowAssert.AssertEquals HttpErrorProtocol, WinHttpErrorMapping.CategoryFromCode(12150)
    XlflowAssert.AssertEquals HttpErrorProtocol, WinHttpErrorMapping.CategoryFromCode(12152)
    XlflowAssert.AssertEquals HttpErrorProtocol, WinHttpErrorMapping.CategoryFromCode(12156)
End Sub

Public Sub Test_NativeErrorMappingDefaultsUnknownCodesToIo()
    XlflowAssert.AssertEquals HttpErrorIo, WinHttpErrorMapping.CategoryFromCode(87)
End Sub
