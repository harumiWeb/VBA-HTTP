Attribute VB_Name = "ReleaseBatchSmoke"
Option Explicit

Public Sub RunBatchSmoke(ByVal releaseWorkbookName As String, ByVal baseUrl As String)
    Dim client As Object
    Dim Urls As New Collection
    Dim result As Object

    Set client = Application.Run("'" & releaseWorkbookName & "'!VBAHttp.CreateClient")
    client.BaseUrl = baseUrl
    Urls.Add "/delay/100"
    Urls.Add "/delay/100"
    Urls.Add "/delay/100"
    Urls.Add "/delay/100"
    Set result = client.GetMany(Urls)

    If result.Count <> 4 Or result.SuccessCount <> 4 Or result.FailureCount <> 0 Or result.CancelledCount <> 0 Then
        Err.Raise vbObjectError + 740, "ReleaseBatchSmoke.RunBatchSmoke", "Release batch smoke returned an unexpected result."
    End If
End Sub
