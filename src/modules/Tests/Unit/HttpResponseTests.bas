Attribute VB_Name = "HttpResponseTests"
Option Explicit

Public Sub Test_Response_4xxIsReturnedAsNonSuccess()
    Dim response As New HttpResponse

    response.Initialize 404, "Not Found"

    XlflowAssert.AssertEquals 404, response.StatusCode
    XlflowAssert.AssertFalse response.IsSuccess
    XlflowAssert.AssertEquals "Not Found", response.ReasonPhrase
End Sub

Public Sub Test_Response_2xxIsSuccess()
    Dim response As New HttpResponse

    response.Initialize 204, "No Content"
    XlflowAssert.AssertTrue response.IsSuccess
End Sub

'@ExpectedError(-2147200494, "HTTP request returned status 500.", "HttpResponse.RaiseForStatus")
Public Sub Test_Response_RaiseForStatusUsesStableError()
    Dim response As New HttpResponse

    response.Initialize 500, "Internal Server Error"
    response.RaiseForStatus
End Sub

Public Sub Test_Response_DecodesDeclaredUtf8()
    Dim headers As New HttpHeaders
    Dim Body As New HttpBody
    Dim response As New HttpResponse
    Dim bytes As Variant

    headers.SetValue "Content-Type", "text/plain; charset=UTF-8"
    bytes = HttpEncoding.EncodeUtf8("caf" & ChrW$(&HE9))
    Body.SetBytes bytes
    response.Initialize 200, "OK", headers, Body

    XlflowAssert.AssertEquals "caf" & ChrW$(&HE9), response.Text
End Sub

Public Sub Test_Response_DefaultsToUtf8()
    Dim Body As New HttpBody
    Dim response As New HttpResponse
    Dim bytes As Variant

    bytes = HttpEncoding.EncodeUtf8(ChrW$(&H6771) & ChrW$(&H4EAC))
    Body.SetBytes bytes
    response.Initialize 200, Body:=Body

    XlflowAssert.AssertEquals ChrW$(&H6771) & ChrW$(&H4EAC), response.Text
End Sub

'@ExpectedError(-2147200503, "Unsupported response charset 'iso-8859-1'.", "HttpResponse.Text")
Public Sub Test_Response_RejectsUnsupportedCharset()
    Dim headers As New HttpHeaders
    Dim Body As New HttpBody
    Dim response As New HttpResponse
    Dim bytes(0 To 0) As Byte

    bytes(0) = 65
    headers.SetValue "Content-Type", "text/plain; charset=iso-8859-1"
    Body.SetBytes bytes
    response.Initialize 200, Headers:=headers, Body:=Body
    Debug.Print response.Text
End Sub
