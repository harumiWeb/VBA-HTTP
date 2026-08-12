Attribute VB_Name = "VBAHttp"
Option Explicit

''' Creates a client with the default buffered WinHTTP transport.
'''
''' Returns:
'''     A new independently configured HttpClient instance.
Public Function CreateClient() As HttpClient
    Set CreateClient = New HttpClient
End Function
