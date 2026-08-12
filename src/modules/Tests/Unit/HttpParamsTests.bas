Attribute VB_Name = "HttpParamsTests"
Option Explicit

Public Sub Test_Params_EncodesUtf8AndReservedCharacters()
    Dim parameters As New HttpParams
    Dim value As String

    value = "caf" & ChrW$(&HE9) & "/" & ChrW$(&H6771) & ChrW$(&H4EAC)
    parameters.Add "q space", value
    XlflowAssert.AssertEquals "q%20space=caf%C3%A9%2F%E6%9D%B1%E4%BA%AC", parameters.ToQueryString()
End Sub

Public Sub Test_Params_PreservesRepeatedNamesAndOrder()
    Dim parameters As New HttpParams

    parameters.Add "tag", "first"
    parameters.Add "tag", "second"
    parameters.Add "active", True

    XlflowAssert.AssertEquals "tag=first&tag=second&active=true", parameters.ToQueryString()
End Sub

Public Sub Test_Params_SetValueReplacesExistingValues()
    Dim parameters As New HttpParams

    parameters.Add "page", 1
    parameters.Add "page", 2
    parameters.SetValue "page", 3

    XlflowAssert.AssertEquals "page=3", parameters.ToQueryString()
End Sub

Public Sub Test_Encoding_RoundTripsSupplementaryUnicode()
    Dim original As String
    Dim bytes As Variant

    original = "A" & ChrW$(&HD83D&) & ChrW$(&HDE80&) & "Z"
    bytes = HttpEncoding.EncodeUtf8(original)

    XlflowAssert.AssertEquals original, HttpEncoding.DecodeUtf8(bytes)
    XlflowAssert.AssertEquals "A%F0%9F%9A%80Z", HttpEncoding.PercentEncode(original)
End Sub

'@ExpectedError(-2147200503, "Response contains malformed UTF-8.", "HttpEncoding.DecodeUtf8")
Public Sub Test_Encoding_RejectsMalformedUtf8()
    Dim bytes(0 To 1) As Byte

    bytes(0) = &HC0
    bytes(1) = &H80
    Call HttpEncoding.DecodeUtf8(bytes)
End Sub
