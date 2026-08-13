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
    XlflowAssert.AssertEquals HttpErrorProtocol, WinHttpErrorMapping.CategoryFromCode(12190)
End Sub

Public Sub Test_NativeErrorMappingDefaultsUnknownCodesToIo()
    XlflowAssert.AssertEquals HttpErrorIo, WinHttpErrorMapping.CategoryFromCode(87)
End Sub

Public Sub Test_NativeOptionCapabilityErrorsIncludeNotSupported()
    XlflowAssert.AssertTrue WinHttpNativeApi.IsOptionCapabilityError(WinHttpNativeApi.WinHttpErrorInvalidOption)
    XlflowAssert.AssertTrue WinHttpNativeApi.IsOptionCapabilityError(WinHttpNativeApi.WinHttpErrorNotSupported)
    XlflowAssert.AssertFalse WinHttpNativeApi.IsOptionCapabilityError(12002)
End Sub

Public Sub Test_NativeUploadLengthEncodingPreservesDwordBits()
    XlflowAssert.AssertEquals 2147483647, WinHttpNativeApi.EncodeUploadTotalLength(CCur(2147483647#))
    XlflowAssert.AssertEquals - 2147483648#, WinHttpNativeApi.EncodeUploadTotalLength(CCur(2147483648#))
    XlflowAssert.AssertEquals - 1, WinHttpNativeApi.EncodeUploadTotalLength(CCur(4294967295#))
    XlflowAssert.AssertEquals WinHttpNativeApi.WinHttpIgnoreRequestTotalLength, WinHttpNativeApi.EncodeUploadTotalLength(CCur(4294967296#))
End Sub

'@ExpectedError(-2147200502, "URL user-info is not permitted.", "WinHttpNativeUrl.Initialize")
Public Sub Test_NativeUrl_RejectsUserInfo()
    Dim Url As New WinHttpNativeUrl

    Url.Initialize "https://user:pass@example.test/"
End Sub

'@ExpectedError(-2147200502, "URL user-info is not permitted.", "WinHttpComTransport.Execute")
Public Sub Test_ComTransport_RejectsUserInfoBeforeNetwork()
    Dim transport As New WinHttpComTransport
    Dim Request As New HttpRequest

    Request.Url = "https://user:pass@example.test/"
    Call transport.Execute(Request)
End Sub
