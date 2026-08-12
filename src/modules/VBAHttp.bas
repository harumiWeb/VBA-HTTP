Attribute VB_Name = "VBAHttp"
Option Explicit

''' Creates a client with the default buffered WinHTTP transport.
'''
''' Returns:
'''     A new independently configured HttpClient instance.
Public Function CreateClient() As HttpClient
    Set CreateClient = New HttpClient
End Function

''' Creates a client using the synchronous native WinHTTP transport.
Public Function CreateNativeClient() As HttpClient
    Dim client As New HttpClient

    Set client.Transport = New WinHttpNativeTransport
    Set CreateNativeClient = client
End Function

''' Creates retry policy configuration for referenced-workbook consumers.
Public Function CreateRetryPolicy() As HttpRetryPolicy
    Set CreateRetryPolicy = New HttpRetryPolicy
End Function

''' Creates per-call reliability options for referenced-workbook consumers.
Public Function CreateExecutionOptions() As HttpExecutionOptions
    Set CreateExecutionOptions = New HttpExecutionOptions
End Function

''' Creates batch options for referenced-workbook consumers.
Public Function CreateBatchOptions() As HttpBatchOptions
    Set CreateBatchOptions = New HttpBatchOptions
End Function

''' Creates a cancellation token for referenced-workbook consumers.
Public Function CreateCancellationToken() As HttpCancellationToken
    Set CreateCancellationToken = New HttpCancellationToken
End Function

''' Creates native HTTP/2 and HTTP/3 negotiation options for consumers.
Public Function CreateProtocolOptions() As HttpProtocolOptions
    Set CreateProtocolOptions = New HttpProtocolOptions
End Function

''' Creates an ordered multipart form for referenced-workbook consumers.
Public Function CreateMultipartForm() As HttpMultipartForm
    Set CreateMultipartForm = New HttpMultipartForm
End Function
