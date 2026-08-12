Attribute VB_Name = "VBAWebBenchmark"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetProcessHandleCount Lib "kernel32" (ByVal processHandle As LongPtr, ByRef handleCount As Long) As Long
Private Declare PtrSafe Function GetProcessMemoryInfo Lib "psapi" (ByVal processHandle As LongPtr, ByRef counters As ProcessMemoryCountersEx, ByVal size As Long) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetProcessHandleCount Lib "kernel32" (ByVal processHandle As Long, ByRef handleCount As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi" (ByVal processHandle As Long, ByRef counters As ProcessMemoryCountersEx, ByVal size As Long) As Long
#End If

Private Type ProcessMemoryCountersEx
    cb As Long
    pageFaultCount As Long
    peakWorkingSetSize As LongPtr
    workingSetSize As LongPtr
    quotaPeakPagedPoolUsage As LongPtr
    quotaPagedPoolUsage As LongPtr
    quotaPeakNonPagedPoolUsage As LongPtr
    quotaNonPagedPoolUsage As LongPtr
    pagefileUsage As LongPtr
    peakPagefileUsage As LongPtr
    privateUsage As LongPtr
End Type

Private Type ProcessSnapshot
    handleCount As Long
    peakWorkingSetBytes As Double
    privateBytes As Double
    workingSetBytes As Double
End Type

Private Const warmupIterations As Long = 5
Private Const latencyIterations As Long = 50
Private Const downloadBytes As Long = 104857600
Private Const timeoutMs As Long = 300000
Private Const upstreamCommit As String = "cefc320acc5372e0b86eed1d20eb3f31b331d598"

Public Sub RunVBAWebBaseline(ByVal baseUrl As String, ByVal outputPath As String)
    ValidateInputs baseUrl, outputPath

    Dim latencyBefore As ProcessSnapshot
    Dim latencyAfter As ProcessSnapshot
    Dim downloadBefore As ProcessSnapshot
    Dim downloadAfter As ProcessSnapshot
    Dim latencyDurations() As Double
    Dim latencyStatus As Long
    Dim downloadStatus As Long
    Dim actualDownloadBytes As Long
    Dim downloadElapsedMs As Double

    WarmUp baseUrl
    latencyBefore = CaptureProcessSnapshot()
    latencyDurations = MeasureLatency(baseUrl, latencyStatus)
    latencyAfter = CaptureProcessSnapshot()
    downloadBefore = CaptureProcessSnapshot()
    MeasureDownload baseUrl, downloadStatus, actualDownloadBytes, downloadElapsedMs
    downloadAfter = CaptureProcessSnapshot()

    WriteResult outputPath, baseUrl, latencyDurations, latencyStatus, downloadStatus, actualDownloadBytes, downloadElapsedMs, latencyBefore, latencyAfter, downloadBefore, downloadAfter
End Sub

Private Sub ValidateInputs(ByVal baseUrl As String, ByVal outputPath As String)
    Const loopbackPrefix As String = "http://127.0.0.1:"
    Dim portText As String

    If Left$(baseUrl, Len(loopbackPrefix)) <> loopbackPrefix Then
        Err.Raise 5, "VBAWebBenchmark.RunVBAWebBaseline", "baseUrl must use the local test server"
    End If
    portText = Mid$(baseUrl, Len(loopbackPrefix) + 1)
    If Len(portText) = 0 Or Not IsNumeric(portText) Then
        Err.Raise 5, "VBAWebBenchmark.RunVBAWebBaseline", "baseUrl must contain a numeric loopback port"
    End If
    If CLng(portText) < 1 Or CLng(portText) > 65535 Or CStr(CLng(portText)) <> portText Then
        Err.Raise 5, "VBAWebBenchmark.RunVBAWebBaseline", "baseUrl contains an invalid loopback port"
    End If
    If Len(outputPath) = 0 Then
        Err.Raise 5, "VBAWebBenchmark.RunVBAWebBaseline", "outputPath is required"
    End If
End Sub

Private Sub WarmUp(ByVal baseUrl As String)
    Dim iteration As Long
    Dim status As Long

    For iteration = 1 To warmupIterations
        status = ExecuteGetStatus(baseUrl, "/status/204")
        If status <> 204 Then
            Err.Raise vbObjectError + 720, "VBAWebBenchmark.WarmUp", "unexpected HTTP status"
        End If
    Next iteration
End Sub

Private Function MeasureLatency(ByVal baseUrl As String, ByRef finalStatus As Long) As Double()
    Dim durations() As Double
    Dim iteration As Long
    Dim started As Currency
    Dim finished As Currency

    ReDim durations(0 To latencyIterations - 1)
    For iteration = 0 To latencyIterations - 1
        started = CounterValue()
        finalStatus = ExecuteGetStatus(baseUrl, "/status/204")
        finished = CounterValue()
        If finalStatus <> 204 Then
            Err.Raise vbObjectError + 721, "VBAWebBenchmark.MeasureLatency", "unexpected HTTP status"
        End If
        durations(iteration) = ElapsedMilliseconds(started, finished)
    Next iteration

    MeasureLatency = durations
End Function

Private Sub MeasureDownload(ByVal baseUrl As String, ByRef status As Long, ByRef actualBytes As Long, ByRef elapsedMs As Double)
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim started As Currency
    Dim finished As Currency

    client.BaseUrl = baseUrl
    client.TimeoutMs = timeoutMs
    request.Method = WebMethod.HttpGet
    request.Resource = "/bytes/" & CStr(downloadBytes)
    request.ResponseFormat = WebFormat.PlainText

    started = CounterValue()
    Set response = client.Execute(request)
    finished = CounterValue()

    status = CLng(response.StatusCode)
    actualBytes = LenB(response.Body)
    elapsedMs = ElapsedMilliseconds(started, finished)
    Set response = Nothing
    Set request = Nothing
    Set client = Nothing

    If status <> 200 Or actualBytes <> downloadBytes Then
        Err.Raise vbObjectError + 722, "VBAWebBenchmark.MeasureDownload", "download response did not match the requested size"
    End If
End Sub

Private Function ExecuteGetStatus(ByVal baseUrl As String, ByVal resource As String) As Long
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse

    client.BaseUrl = baseUrl
    client.TimeoutMs = timeoutMs
    request.Method = WebMethod.HttpGet
    request.Resource = resource
    request.ResponseFormat = WebFormat.PlainText
    Set response = client.Execute(request)
    ExecuteGetStatus = CLng(response.StatusCode)
    Set response = Nothing
    Set request = Nothing
    Set client = Nothing
End Function

Private Function CounterValue() As Currency
    If QueryPerformanceCounter(CounterValue) = 0 Then
        Err.Raise vbObjectError + 723, "VBAWebBenchmark.CounterValue", "QueryPerformanceCounter failed"
    End If
End Function

Private Function ElapsedMilliseconds(ByVal started As Currency, ByVal finished As Currency) As Double
    Static frequency As Currency

    If frequency = 0 Then
        If QueryPerformanceFrequency(frequency) = 0 Then
            Err.Raise vbObjectError + 724, "VBAWebBenchmark.ElapsedMilliseconds", "QueryPerformanceFrequency failed"
        End If
    End If
    ElapsedMilliseconds = (CDbl(finished) - CDbl(started)) * 1000# / CDbl(frequency)
End Function

Private Function CaptureProcessSnapshot() As ProcessSnapshot
    Dim counters As ProcessMemoryCountersEx
    Dim snapshot As ProcessSnapshot
    Dim processHandle As LongPtr

    counters.cb = LenB(counters)
    processHandle = GetCurrentProcess()
    If GetProcessMemoryInfo(processHandle, counters, counters.cb) = 0 Then
        Err.Raise vbObjectError + 725, "VBAWebBenchmark.CaptureProcessSnapshot", "GetProcessMemoryInfo failed"
    End If
    If GetProcessHandleCount(processHandle, snapshot.handleCount) = 0 Then
        Err.Raise vbObjectError + 726, "VBAWebBenchmark.CaptureProcessSnapshot", "GetProcessHandleCount failed"
    End If

    snapshot.workingSetBytes = PointerSizedValue(counters.workingSetSize)
    snapshot.peakWorkingSetBytes = PointerSizedValue(counters.peakWorkingSetSize)
    snapshot.privateBytes = PointerSizedValue(counters.privateUsage)
    CaptureProcessSnapshot = snapshot
End Function

Private Function PointerSizedValue(ByVal value As LongPtr) As Double
    #If Win64 Then
    PointerSizedValue = CDbl(value)
    #Else
    If value < 0 Then
        PointerSizedValue = 4294967296# + CDbl(value)
    Else
        PointerSizedValue = CDbl(value)
    End If
    #End If
End Function

Private Sub WriteResult(ByVal outputPath As String, ByVal baseUrl As String, ByRef durations() As Double, ByVal latencyStatus As Long, ByVal downloadStatus As Long, ByVal actualDownloadBytes As Long, ByVal downloadElapsedMs As Double, ByRef latencyBefore As ProcessSnapshot, ByRef latencyAfter As ProcessSnapshot, ByRef downloadBefore As ProcessSnapshot, ByRef downloadAfter As ProcessSnapshot)
    Dim fileSystem As Object
    Dim output As Object
    Dim totalLatencyMs As Double
    Dim meanLatencyMs As Double
    Dim minimumLatencyMs As Double
    Dim maximumLatencyMs As Double
    Dim requestsPerSecond As Double
    Dim downloadMiBPerSecond As Double

    SummarizeDurations durations, totalLatencyMs, meanLatencyMs, minimumLatencyMs, maximumLatencyMs
    requestsPerSecond = latencyIterations * 1000# / totalLatencyMs
    downloadMiBPerSecond = (actualDownloadBytes / 1048576#) / (downloadElapsedMs / 1000#)

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FileExists(outputPath) Then
        Err.Raise vbObjectError + 727, "VBAWebBenchmark.WriteResult", "outputPath already exists"
    End If
    Set output = fileSystem.CreateTextFile(outputPath, False, False)
    If output Is Nothing Then
        Err.Raise vbObjectError + 728, "VBAWebBenchmark.WriteResult", "could not create benchmark result"
    End If

    output.WriteLine "{"
    output.WriteLine "  ""schema_version"": 1,"
    output.WriteLine "  ""benchmark"": ""http-client"","
    output.WriteLine "  ""implementation"": {""name"": ""VBA-Web"", ""version"": ""4.1.6"", ""source_commit"": """ & upstreamCommit & """},"
    output.WriteLine "  ""environment"": {""office_bitness"": """ & OfficeBitness() & """, ""excel_version"": """ & JsonEscape(Application.Version) & """, ""platform"": ""Windows""},"
    output.WriteLine "  ""server"": {""base_url"": """ & JsonEscape(baseUrl) & """, ""external_network"": false},"
    output.WriteLine "  ""parameters"": {""warmup_iterations"": " & CStr(warmupIterations) & ", ""latency_iterations"": " & CStr(latencyIterations) & ", ""download_bytes"": " & CStr(downloadBytes) & ", ""timeouts_ms"": {""resolve"": " & CStr(timeoutMs) & ", ""connect"": " & CStr(timeoutMs) & ", ""send"": " & CStr(timeoutMs) & ", ""receive"": " & CStr(timeoutMs) & "}},"
    output.WriteLine "  ""results"": ["
    output.WriteLine "    {""scenario"": ""sequential_get"", ""method"": ""GET"", ""path"": ""/status/204"", ""iterations"": " & CStr(latencyIterations) & ", ""status"": " & CStr(latencyStatus) & ", ""elapsed_ms"": " & JsonNumber(totalLatencyMs) & ", ""mean_ms"": " & JsonNumber(meanLatencyMs) & ", ""min_ms"": " & JsonNumber(minimumLatencyMs) & ", ""max_ms"": " & JsonNumber(maximumLatencyMs) & ", ""requests_per_second"": " & JsonNumber(requestsPerSecond) & ", ""process_before"": " & SnapshotJson(latencyBefore) & ", ""process_after"": " & SnapshotJson(latencyAfter) & "},"
    output.WriteLine "    {""scenario"": ""buffered_download"", ""method"": ""GET"", ""path"": ""/bytes/" & CStr(downloadBytes) & """, ""iterations"": 1, ""status"": " & CStr(downloadStatus) & ", ""elapsed_ms"": " & JsonNumber(downloadElapsedMs) & ", ""bytes"": " & CStr(actualDownloadBytes) & ", ""throughput_mib_s"": " & JsonNumber(downloadMiBPerSecond) & ", ""process_before"": " & SnapshotJson(downloadBefore) & ", ""process_after"": " & SnapshotJson(downloadAfter) & "}"
    output.WriteLine "  ]"
    output.WriteLine "}"
    output.Close
    Set output = Nothing
    Set fileSystem = Nothing
End Sub

Private Sub SummarizeDurations(ByRef durations() As Double, ByRef totalMs As Double, ByRef meanMs As Double, ByRef minimumMs As Double, ByRef maximumMs As Double)
    Dim index As Long

    minimumMs = durations(LBound(durations))
    maximumMs = minimumMs
    For index = LBound(durations) To UBound(durations)
        totalMs = totalMs + durations(index)
        If durations(index) < minimumMs Then minimumMs = durations(index)
        If durations(index) > maximumMs Then maximumMs = durations(index)
    Next index
    meanMs = totalMs / (UBound(durations) - LBound(durations) + 1)
End Sub

Private Function SnapshotJson(ByRef snapshot As ProcessSnapshot) As String
    SnapshotJson = "{""handles"": " & CStr(snapshot.handleCount) & ", ""working_set_bytes"": " & JsonNumber(snapshot.workingSetBytes) & ", ""peak_working_set_bytes"": " & JsonNumber(snapshot.peakWorkingSetBytes) & ", ""private_bytes"": " & JsonNumber(snapshot.privateBytes) & "}"
End Function

Private Function JsonNumber(ByVal value As Double) As String
    JsonNumber = Trim$(Str$(Round(value, 3)))
    If Left$(JsonNumber, 1) = "." Then
        JsonNumber = "0" & JsonNumber
    ElseIf Left$(JsonNumber, 2) = "-." Then
        JsonNumber = "-0" & Mid$(JsonNumber, 2)
    End If
End Function

Private Function JsonEscape(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", "\""")
    value = Replace(value, vbCr, "\r")
    value = Replace(value, vbLf, "\n")
    value = Replace(value, vbTab, "\t")
    JsonEscape = value
End Function

Private Function OfficeBitness() As String
    #If Win64 Then
    OfficeBitness = "x64"
    #Else
    OfficeBitness = "x86"
    #End If
End Function
