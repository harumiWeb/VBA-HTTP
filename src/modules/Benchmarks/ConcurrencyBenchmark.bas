Attribute VB_Name = "ConcurrencyBenchmark"
Option Explicit

Private Const BenchmarkRequests As Long = 100
Private Const DelayMilliseconds As Long = 100
Private Const ConcurrentLimit As Long = 16

Public Sub RunConcurrencyBaseline(ByVal baseUrl As String, ByVal outputPath As String)
    Dim client As New HttpClient
    Dim Options As New HttpBatchOptions
    Dim sequentialMilliseconds As Double
    Dim concurrentMilliseconds As Double
    Dim maxInFlight As Long

    ValidateInputs baseUrl, outputPath
    client.BaseUrl = baseUrl
    Options.YieldToHost = False
    Options.PollIntervalMilliseconds = 1

    WarmUp client
    Options.MaxConcurrency = 1
    sequentialMilliseconds = MeasureBatch(client, Options)

    Call client.PostResponse("/__admin/reset")
    Options.MaxConcurrency = ConcurrentLimit
    concurrentMilliseconds = MeasureBatch(client, Options)
    maxInFlight = CLng(client.GetResponse("/__admin/stats").Headers.GetValue("X-Max-In-Flight"))

    WriteResult outputPath, sequentialMilliseconds, concurrentMilliseconds, maxInFlight
End Sub

Private Sub ValidateInputs(ByVal baseUrl As String, ByVal outputPath As String)
    Const loopbackPrefix As String = "http://127.0.0.1:"
    Dim portText As String

    If Left$(baseUrl, Len(loopbackPrefix)) <> loopbackPrefix Then
        Err.Raise 5, "ConcurrencyBenchmark.RunConcurrencyBaseline", "baseUrl must use the local test server"
    End If
    portText = Mid$(baseUrl, Len(loopbackPrefix) + 1)
    If Len(portText) = 0 Or Not IsNumeric(portText) Then
        Err.Raise 5, "ConcurrencyBenchmark.RunConcurrencyBaseline", "baseUrl must contain a numeric loopback port"
    End If
    If CLng(portText) < 1 Or CLng(portText) > 65535 Or CStr(CLng(portText)) <> portText Then
        Err.Raise 5, "ConcurrencyBenchmark.RunConcurrencyBaseline", "baseUrl contains an invalid loopback port"
    End If
    If Len(outputPath) = 0 Then
        Err.Raise 5, "ConcurrencyBenchmark.RunConcurrencyBaseline", "outputPath is required"
    End If
End Sub

Private Sub WarmUp(ByVal client As HttpClient)
    Dim index As Long
    Dim response As HttpResponse

    For index = 1 To 5
        Set response = client.GetResponse("/status/204")
        If response.StatusCode <> 204 Then Err.Raise vbObjectError + 730, "ConcurrencyBenchmark.WarmUp", "Warmup returned an unexpected status."
    Next index
End Sub

Private Function MeasureBatch(ByVal client As HttpClient, ByVal Options As HttpBatchOptions) As Double
    Dim result As HttpBatchResult
    Dim started As Currency

    started = HttpTiming.CounterValue()
    Set result = client.ExecuteMany(CreateRequests(), Options)
    MeasureBatch = HttpTiming.ElapsedMilliseconds(started, HttpTiming.CounterValue())
    If result.Count <> BenchmarkRequests Or result.SuccessCount <> BenchmarkRequests Then
        Err.Raise vbObjectError + 731, "ConcurrencyBenchmark.MeasureBatch", "Batch did not complete every request successfully."
    End If
End Function

Private Function CreateRequests() As Collection
    Dim Requests As New Collection
    Dim Request As HttpRequest
    Dim index As Long

    For index = 1 To BenchmarkRequests
        Set Request = New HttpRequest
        Request.Method = "GET"
        Request.Url = "/delay/" & CStr(DelayMilliseconds)
        Requests.Add Request
    Next index
    Set CreateRequests = Requests
End Function

Private Sub WriteResult(ByVal outputPath As String, ByVal sequentialMilliseconds As Double, ByVal concurrentMilliseconds As Double, ByVal maxInFlight As Long)
    Dim fileSystem As Object
    Dim output As Object
    Dim speedup As Double

    speedup = sequentialMilliseconds / concurrentMilliseconds
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FileExists(outputPath) Then
        Err.Raise vbObjectError + 732, "ConcurrencyBenchmark.WriteResult", "outputPath already exists"
    End If
    ' The external runner confines outputPath to benchmarks/results and refuses overwrite.
    Set output = fileSystem.CreateTextFile(outputPath, False, False) ' xlflow:disable-line VBA245
    If output Is Nothing Then Err.Raise vbObjectError + 733, "ConcurrencyBenchmark.WriteResult", "could not create benchmark result"
    output.WriteLine "{" ' xlflow:disable-line VBA202
    output.WriteLine "  ""schema_version"": 1,"
    output.WriteLine "  ""benchmark"": ""bounded-concurrency"","
    output.WriteLine "  ""environment"": {""office_bitness"": """ & OfficeBitness() & """, ""excel_version"": """ & Application.Version & """, ""platform"": ""Windows""},"
    output.WriteLine "  ""parameters"": {""requests"": " & CStr(BenchmarkRequests) & ", ""delay_ms"": " & CStr(DelayMilliseconds) & ", ""sequential_concurrency"": 1, ""concurrent_concurrency"": " & CStr(ConcurrentLimit) & "},"
    output.WriteLine "  ""results"": {""sequential_elapsed_ms"": " & JsonNumber(sequentialMilliseconds) & ", ""concurrent_elapsed_ms"": " & JsonNumber(concurrentMilliseconds) & ", ""speedup"": " & JsonNumber(speedup) & ", ""server_max_in_flight"": " & CStr(maxInFlight) & "}"
    output.WriteLine "}"
    output.Close
    Set output = Nothing
    Set fileSystem = Nothing
End Sub

Private Function JsonNumber(ByVal value As Double) As String
    JsonNumber = Trim$(Str$(Round(value, 3)))
    If Left$(JsonNumber, 1) = "." Then JsonNumber = "0" & JsonNumber
End Function

Private Function OfficeBitness() As String
    #If Win64 Then
    OfficeBitness = "x64"
    #Else
    OfficeBitness = "x86"
    #End If
End Function
