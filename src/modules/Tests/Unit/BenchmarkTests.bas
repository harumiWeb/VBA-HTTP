Attribute VB_Name = "BenchmarkTests"
Option Explicit

'@ExpectedError(5, "baseUrl must use the local test server", "RawWinHttpBenchmark.RunRawBaseline")
Public Sub Test_RawBenchmark_RejectsExternalNetwork()
    RawWinHttpBenchmark.RunRawBaseline "https://example.com", "unused.json"
End Sub

'@ExpectedError(5, "baseUrl contains an invalid loopback port", "RawWinHttpBenchmark.RunRawBaseline")
Public Sub Test_RawBenchmark_RejectsInvalidLoopbackPort()
    RawWinHttpBenchmark.RunRawBaseline "http://127.0.0.1:70000", "unused.json"
End Sub
