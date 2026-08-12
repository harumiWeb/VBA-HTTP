Attribute VB_Name = "ConsumerDiagnostics"
Option Explicit

' Diagnostics are opt-in and redact credential/cookie-bearing headers.
Public Sub CaptureSafeOperationEvents()
    Dim client As HttpClient
    Dim diagnostics As HttpDiagnostics
    Dim response As HttpResponse

    Set diagnostics = VBAHttp.CreateDiagnostics()
    diagnostics.Enabled = True
    Set client = VBAHttp.CreateClient()
    Set client.Diagnostics = diagnostics

    Set response = client.GetResponse("http://127.0.0.1:8080/status/204")
    response.RaiseForStatus
    Debug.Print diagnostics.ToJson
End Sub
