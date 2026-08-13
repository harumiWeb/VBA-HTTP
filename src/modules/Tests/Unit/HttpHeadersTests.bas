Attribute VB_Name = "HttpHeadersTests"
Option Explicit

Public Sub Test_Headers_LookupIsCaseInsensitive()
    Dim headers As New HttpHeaders

    headers.SetValue "Content-Type", "application/json"
    XlflowAssert.AssertEquals "application/json", headers.GetValue("content-type")
    XlflowAssert.AssertTrue headers.Contains("CONTENT-TYPE")
End Sub

Public Sub Test_Headers_AddPreservesRepeatedValues()
    Dim headers As New HttpHeaders
    Dim values As Variant

    headers.Add "Set-Cookie", "a=1"
    headers.Add "set-cookie", "b=2"
    values = headers.GetValues("SET-cookie")

    XlflowAssert.AssertArrayEquals Array("a=1", "b=2"), values
End Sub

Public Sub Test_Headers_SetValueReplacesRepeatedValues()
    Dim headers As New HttpHeaders

    headers.Add "Accept", "text/plain"
    headers.Add "Accept", "application/json"
    headers.SetValue "accept", "application/octet-stream"

    XlflowAssert.AssertEquals 1, headers.Count
    XlflowAssert.AssertEquals "application/octet-stream", headers.GetValue("ACCEPT")
End Sub

Public Sub Test_Headers_CloneIsIndependent()
    Dim headers As New HttpHeaders
    Dim copy As HttpHeaders

    headers.SetValue "X-Test", "before"
    Set copy = headers.Clone()
    copy.SetValue "X-Test", "after"

    XlflowAssert.AssertEquals "before", headers.GetValue("X-Test")
    XlflowAssert.AssertEquals "after", copy.GetValue("X-Test")
End Sub

'@ExpectedError(-2147200503, "Header name contains an invalid character.", "HttpHeaders")
Public Sub Test_Headers_RejectsColonInName()
    Dim headers As New HttpHeaders
    headers.Add "Bad:Name", "value"
End Sub

'@ExpectedError(-2147200503, "Header name contains an invalid character.", "HttpHeaders")
Public Sub Test_Headers_RejectsUnicodeNameOutsideAsciiToken()
    Dim headers As New HttpHeaders

    headers.Add ChrW$(-223), "value"
End Sub

'@ExpectedError(-2147200503, "Header value cannot contain CR or LF.", "HttpHeaders")
Public Sub Test_Headers_RejectsLineBreakInValue()
    Dim headers As New HttpHeaders
    headers.Add "X-Test", "safe" & vbCrLf & "Injected: true"
End Sub

'@ExpectedError(-2147200503, "Header value cannot contain control characters.", "HttpHeaders")
Public Sub Test_Headers_RejectsControlCharactersInValue()
    Dim headers As New HttpHeaders

    headers.Add "X-Test", "safe" & ChrW$(0) & "invalid"
End Sub

Public Sub Test_Headers_AllowsHorizontalTabAndUnicodeInValue()
    Dim headers As New HttpHeaders

    headers.Add "X-Test", "値" & ChrW$(9) & "ok"
    XlflowAssert.AssertEquals "値" & ChrW$(9) & "ok", headers.GetValue("X-Test")
End Sub
