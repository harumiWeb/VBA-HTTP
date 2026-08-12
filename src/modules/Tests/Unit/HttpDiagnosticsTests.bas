Attribute VB_Name = "HttpDiagnosticsTests"
Option Explicit

Public Sub Test_Diagnostics_IsDisabledAndBoundedByDefault()
    Dim diagnostics As New HttpDiagnostics

    XlflowAssert.AssertFalse diagnostics.Enabled
    XlflowAssert.AssertEquals 100, diagnostics.MaxEvents
    XlflowAssert.AssertEquals 0, diagnostics.Count
End Sub

Public Sub Test_Diagnostics_CapturesRedactedResponseEvent()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim diagnostics As New HttpDiagnostics
    Dim Request As New HttpRequest
    Dim eventItem As HttpDiagnosticEvent
    Dim json As String

    configuredResponse.Initialize 204, "No Content"
    configuredResponse.Headers.SetValue "X-Trace", "safe-value"
    configuredResponse.Headers.SetValue "Set-Cookie", "session=secret-response"
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    diagnostics.Enabled = True
    Set client.Diagnostics = diagnostics
    Request.Method = "GET"
    Request.Url = "https://user:password@example.test/items?token=query-secret#fragment"
    Request.Headers.SetValue "Authorization", "Bearer request-secret"
    Request.Headers.SetValue "X-Trace", "safe-request"

    Call client.Execute(Request)

    XlflowAssert.AssertEquals 1, diagnostics.Count
    Set eventItem = diagnostics.ItemAt(1)
    XlflowAssert.AssertEquals "execute", eventItem.Operation
    XlflowAssert.AssertEquals "GET", eventItem.Method
    XlflowAssert.AssertEquals "https://example.test/items", eventItem.Target
    XlflowAssert.AssertEquals 204, eventItem.StatusCode
    XlflowAssert.AssertEquals "[REDACTED]", eventItem.RequestHeaders.GetValue("Authorization")
    XlflowAssert.AssertEquals "[REDACTED]", eventItem.ResponseHeaders.GetValue("Set-Cookie")
    XlflowAssert.AssertEquals "safe-value", eventItem.ResponseHeaders.GetValue("X-Trace")

    json = diagnostics.ToJson()
    XlflowAssert.AssertTrue InStr(1, json, "[REDACTED]", vbBinaryCompare) > 0
    XlflowAssert.AssertFalse InStr(1, json, "request-secret", vbBinaryCompare) > 0
    XlflowAssert.AssertFalse InStr(1, json, "session=secret-response", vbBinaryCompare) > 0
    XlflowAssert.AssertFalse InStr(1, json, "query-secret", vbBinaryCompare) > 0
    XlflowAssert.AssertTrue InStr(1, json, "schema_version", vbBinaryCompare) > 0
End Sub

Public Sub Test_Diagnostics_CapturesStableFailureWithoutDescription()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim diagnostics As New HttpDiagnostics
    Dim Request As New HttpRequest
    Dim observedNumber As Long
    Dim json As String

    transport.SetFailure HttpErrTimeout, "secret backend failure description"
    Set client.Transport = transport
    diagnostics.Enabled = True
    Set client.Diagnostics = diagnostics
    Request.Url = "https://example.test/slow?secret=query-value"

    On Error Resume Next
    Call client.Execute(Request)
    observedNumber = Err.Number
    Err.Clear
    On Error GoTo 0

    XlflowAssert.AssertEquals HttpErrTimeout, observedNumber
    XlflowAssert.AssertEquals 1, diagnostics.Count
    XlflowAssert.AssertEquals HttpErrTimeout, diagnostics.LastEvent.ErrorNumber
    XlflowAssert.AssertEquals HttpErrorTimeout, diagnostics.LastEvent.ErrorCategory
    json = diagnostics.ToJson()
    XlflowAssert.AssertFalse InStr(1, json, "secret backend failure description", vbBinaryCompare) > 0
    XlflowAssert.AssertFalse InStr(1, json, "query-value", vbBinaryCompare) > 0
End Sub

Public Sub Test_Diagnostics_DropsOldestEventAtMaxEvents()
    Dim client As New HttpClient
    Dim transport As New MockHttpTransport
    Dim configuredResponse As New HttpResponse
    Dim diagnostics As New HttpDiagnostics
    Dim index As Long
    Dim Request As New HttpRequest

    configuredResponse.Initialize 204
    transport.SetResponse configuredResponse
    Set client.Transport = transport
    diagnostics.Enabled = True
    diagnostics.MaxEvents = 2
    Set client.Diagnostics = diagnostics

    For index = 1 To 3
        Request.Url = "https://example.test/items/" & CStr(index)
        Call client.Execute(Request)
    Next index

    XlflowAssert.AssertEquals 2, diagnostics.Count
    XlflowAssert.AssertEquals 2, diagnostics.ItemAt(1).Sequence
    XlflowAssert.AssertEquals "https://example.test/items/2", diagnostics.ItemAt(1).Target
    XlflowAssert.AssertEquals 3, diagnostics.ItemAt(2).Sequence
End Sub

Public Sub Test_Diagnostics_FactoryReturnsCollector()
    Dim diagnostics As HttpDiagnostics

    Set diagnostics = VBAHttp.CreateDiagnostics()
    XlflowAssert.AssertIsNotNothing diagnostics
    diagnostics.Enabled = True
    XlflowAssert.AssertTrue diagnostics.Enabled
End Sub

Public Sub Test_Diagnostics_JsonQuotesControlCharacters()
    Dim value As String
    Dim expected As String

    value = "quote" & Chr$(34) & "slash" & Chr$(92) & vbTab
    expected = Chr$(34) & "quote" & Chr$(92) & Chr$(34) & "slash" & Chr$(92) & Chr$(92) & Chr$(92) & "t" & Chr$(34)
    XlflowAssert.AssertEquals expected, HttpJson.Quote(value)
End Sub
