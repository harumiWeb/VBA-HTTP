Attribute VB_Name = "WinHttpNativeApi"
Option Explicit

Public Const WinHttpAccessTypeDefaultProxy As Long = 0
Public Const WinHttpFlagSecure As Long = &H800000
Public Const WinHttpOptionRedirectPolicy As Long = 88
Public Const WinHttpOptionMaxAutomaticRedirects As Long = 89
Public Const WinHttpOptionHttpProtocolUsed As Long = 134
Public Const WinHttpProtocolFlagHttp2 As Long = 1
Public Const WinHttpProtocolFlagHttp3 As Long = 2
Public Const WinHttpQueryStatusCode As Long = 19
Public Const WinHttpQueryStatusText As Long = 20
Public Const WinHttpQueryRawHeadersCrlf As Long = 22
Public Const WinHttpQueryFlagNumber As Long = &H20000000
Public Const WinHttpAddRequestHeader As Long = &H20000000
Public Const WinHttpErrorInsufficientBuffer As Long = 122
Public Const WinHttpErrorHeaderNotFound As Long = 12150
Public Const WinHttpErrorInvalidOption As Long = 12009

#If VBA7 Then
Private Declare PtrSafe Function WinHttpOpen Lib "winhttp.dll" (ByVal pwszUserAgent As LongPtr, ByVal dwAccessType As Long, ByVal pwszProxyName As LongPtr, ByVal pwszProxyBypass As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp.dll" (ByVal hInternet As LongPtr) As Long
Private Declare PtrSafe Function WinHttpConnect Lib "winhttp.dll" (ByVal hSession As LongPtr, ByVal pswzServerName As LongPtr, ByVal nServerPort As Long, ByVal dwReserved As Long) As LongPtr
Private Declare PtrSafe Function WinHttpOpenRequest Lib "winhttp.dll" (ByVal hConnect As LongPtr, ByVal pwszVerb As LongPtr, ByVal pwszObjectName As LongPtr, ByVal pwszVersion As LongPtr, ByVal pwszReferrer As LongPtr, ByVal ppwszAcceptTypes As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpAddRequestHeaders Lib "winhttp.dll" (ByVal hRequest As LongPtr, ByVal lpszHeaders As LongPtr, ByVal dwHeadersLength As Long, ByVal dwModifiers As Long) As Long
Private Declare PtrSafe Function WinHttpSetTimeouts Lib "winhttp.dll" (ByVal hInternet As LongPtr, ByVal nResolveTimeout As Long, ByVal nConnectTimeout As Long, ByVal nSendTimeout As Long, ByVal nReceiveTimeout As Long) As Long
Private Declare PtrSafe Function WinHttpSendRequest Lib "winhttp.dll" (ByVal hRequest As LongPtr, ByVal lpszHeaders As LongPtr, ByVal dwHeadersLength As Long, ByVal lpOptional As LongPtr, ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, ByVal dwContext As LongPtr) As Long
Private Declare PtrSafe Function WinHttpReceiveResponse Lib "winhttp.dll" (ByVal hRequest As LongPtr, ByVal lpReserved As LongPtr) As Long
Private Declare PtrSafe Function WinHttpQueryHeaders Lib "winhttp.dll" (ByVal hRequest As LongPtr, ByVal dwInfoLevel As Long, ByVal pwszName As LongPtr, ByVal lpBuffer As LongPtr, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Long
Private Declare PtrSafe Function WinHttpQueryDataAvailable Lib "winhttp.dll" (ByVal hRequest As LongPtr, ByRef lpdwNumberOfBytesAvailable As Long) As Long
Private Declare PtrSafe Function WinHttpReadData Lib "winhttp.dll" (ByVal hRequest As LongPtr, ByRef lpBuffer As Any, ByVal dwNumberOfBytesToRead As Long, ByRef lpdwNumberOfBytesRead As Long) As Long
Private Declare PtrSafe Function WinHttpQueryOption Lib "winhttp.dll" (ByVal hInternet As LongPtr, ByVal dwOption As Long, ByVal lpBuffer As LongPtr, ByRef lpdwBufferLength As Long) As Long
Private Declare PtrSafe Function WinHttpSetOption Lib "winhttp.dll" (ByVal hInternet As LongPtr, ByVal dwOption As Long, ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetProcessHandleCount Lib "kernel32" (ByVal hProcess As LongPtr, ByRef pdwHandleCount As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Function WinHttpOpen Lib "winhttp.dll" (ByVal pwszUserAgent As Long, ByVal dwAccessType As Long, ByVal pwszProxyName As Long, ByVal pwszProxyBypass As Long, ByVal dwFlags As Long) As Long
Private Declare Function WinHttpCloseHandle Lib "winhttp.dll" (ByVal hInternet As Long) As Long
Private Declare Function WinHttpConnect Lib "winhttp.dll" (ByVal hSession As Long, ByVal pswzServerName As Long, ByVal nServerPort As Long, ByVal dwReserved As Long) As Long
Private Declare Function WinHttpOpenRequest Lib "winhttp.dll" (ByVal hConnect As Long, ByVal pwszVerb As Long, ByVal pwszObjectName As Long, ByVal pwszVersion As Long, ByVal pwszReferrer As Long, ByVal ppwszAcceptTypes As Long, ByVal dwFlags As Long) As Long
Private Declare Function WinHttpAddRequestHeaders Lib "winhttp.dll" (ByVal hRequest As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal dwModifiers As Long) As Long
Private Declare Function WinHttpSetTimeouts Lib "winhttp.dll" (ByVal hInternet As Long, ByVal nResolveTimeout As Long, ByVal nConnectTimeout As Long, ByVal nSendTimeout As Long, ByVal nReceiveTimeout As Long) As Long
Private Declare Function WinHttpSendRequest Lib "winhttp.dll" (ByVal hRequest As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal lpOptional As Long, ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, ByVal dwContext As Long) As Long
Private Declare Function WinHttpReceiveResponse Lib "winhttp.dll" (ByVal hRequest As Long, ByVal lpReserved As Long) As Long
Private Declare Function WinHttpQueryHeaders Lib "winhttp.dll" (ByVal hRequest As Long, ByVal dwInfoLevel As Long, ByVal pwszName As Long, ByVal lpBuffer As Long, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Long
Private Declare Function WinHttpQueryDataAvailable Lib "winhttp.dll" (ByVal hRequest As Long, ByRef lpdwNumberOfBytesAvailable As Long) As Long
Private Declare Function WinHttpReadData Lib "winhttp.dll" (ByVal hRequest As Long, ByRef lpBuffer As Any, ByVal dwNumberOfBytesToRead As Long, ByRef lpdwNumberOfBytesRead As Long) As Long
Private Declare Function WinHttpQueryOption Lib "winhttp.dll" (ByVal hInternet As Long, ByVal dwOption As Long, ByVal lpBuffer As Long, ByRef lpdwBufferLength As Long) As Long
Private Declare Function WinHttpSetOption Lib "winhttp.dll" (ByVal hInternet As Long, ByVal dwOption As Long, ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetProcessHandleCount Lib "kernel32" (ByVal hProcess As Long, ByRef pdwHandleCount As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If

#If VBA7 Then
Public Function OpenSession(ByVal UserAgent As String) As LongPtr
    OpenSession = WinHttpOpen(StrPtr(UserAgent), WinHttpAccessTypeDefaultProxy, 0, 0, 0)
End Function

Public Function Connect(ByVal SessionHandle As LongPtr, ByVal ServerName As String, ByVal ServerPort As Long) As LongPtr
    Connect = WinHttpConnect(SessionHandle, StrPtr(ServerName), ServerPort, 0)
End Function

Public Function OpenRequest(ByVal ConnectionHandle As LongPtr, ByVal Method As String, ByVal ObjectName As String, ByVal Flags As Long) As LongPtr
    OpenRequest = WinHttpOpenRequest(ConnectionHandle, StrPtr(Method), StrPtr(ObjectName), 0, 0, 0, Flags)
End Function

Public Function CloseHandle(ByVal HandleValue As LongPtr) As Boolean
    If HandleValue = 0 Then
        CloseHandle = True
    Else
        CloseHandle = (WinHttpCloseHandle(HandleValue) <> 0)
    End If
End Function

Public Function SetTimeouts(ByVal HandleValue As LongPtr, ByVal ResolveMilliseconds As Long, ByVal ConnectMilliseconds As Long, ByVal SendMilliseconds As Long, ByVal ReceiveMilliseconds As Long) As Boolean
    SetTimeouts = (WinHttpSetTimeouts(HandleValue, ResolveMilliseconds, ConnectMilliseconds, SendMilliseconds, ReceiveMilliseconds) <> 0)
End Function

Public Function AddRequestHeaders(ByVal HandleValue As LongPtr, ByVal Headers As String) As Boolean
    If Len(Headers) = 0 Then
        AddRequestHeaders = True
    Else
        AddRequestHeaders = (WinHttpAddRequestHeaders(HandleValue, StrPtr(Headers), Len(Headers), WinHttpAddRequestHeader) <> 0)
    End If
End Function

Public Function SendBody(ByVal HandleValue As LongPtr, ByVal Body As Variant) As Boolean
    Dim bytes() As Byte
    Dim lowerBound As Long
    Dim byteCount As Long

    If Not HttpEncoding.HasByteArray(Body) Then
        SendBody = (WinHttpSendRequest(HandleValue, 0, 0, 0, 0, 0, 0) <> 0)
        Exit Function
    End If
    bytes = Body
    lowerBound = LBound(bytes) ' xlflow:disable-line VBA227
    byteCount = UBound(bytes) - lowerBound + 1 ' xlflow:disable-line VBA227
    SendBody = (WinHttpSendRequest(HandleValue, 0, 0, VarPtr(bytes(lowerBound)), byteCount, byteCount, 0) <> 0) ' xlflow:disable-line VBA227
End Function

Public Function ReceiveResponse(ByVal HandleValue As LongPtr) As Boolean
    ReceiveResponse = (WinHttpReceiveResponse(HandleValue, 0) <> 0)
End Function

Public Function QueryHeaderString(ByVal HandleValue As LongPtr, ByVal InfoLevel As Long, ByRef Value As String) As Boolean
    Dim requiredBytes As Long
    Dim headerIndex As Long
    Dim buffer As String
    Dim errorCode As Long

    requiredBytes = 0
    headerIndex = 0
    If WinHttpQueryHeaders(HandleValue, InfoLevel, 0, 0, requiredBytes, headerIndex) <> 0 Then
        Value = ""
        QueryHeaderString = True
        Exit Function
    End If
    errorCode = GetLastError()
    If errorCode = WinHttpErrorHeaderNotFound Then
        Value = ""
        QueryHeaderString = True
        Exit Function
    End If
    If errorCode <> WinHttpErrorInsufficientBuffer Or requiredBytes <= 0 Or (requiredBytes Mod 2) <> 0 Then Exit Function

    buffer = String$(requiredBytes \ 2, vbNullChar)
    If WinHttpQueryHeaders(HandleValue, InfoLevel, 0, StrPtr(buffer), requiredBytes, headerIndex) = 0 Then Exit Function
    Value = RemoveTrailingNull(buffer)
    QueryHeaderString = True
End Function

Public Function QueryDataAvailable(ByVal HandleValue As LongPtr, ByRef ByteCount As Long) As Boolean
    QueryDataAvailable = (WinHttpQueryDataAvailable(HandleValue, ByteCount) <> 0)
End Function

Public Function ReadData(ByVal HandleValue As LongPtr, ByRef Buffer() As Byte, ByVal RequestedBytes As Long, ByRef ReadBytes As Long) As Boolean
    ReadData = (WinHttpReadData(HandleValue, Buffer(0), RequestedBytes, ReadBytes) <> 0)
End Function

Public Function QueryProtocolUsed(ByVal HandleValue As LongPtr, ByRef ProtocolFlags As Long) As Boolean
    Dim bufferLength As Long

    ProtocolFlags = 0
    bufferLength = 4
    QueryProtocolUsed = (WinHttpQueryOption(HandleValue, WinHttpOptionHttpProtocolUsed, VarPtr(ProtocolFlags), bufferLength) <> 0)
End Function

Public Function SetOptionLong(ByVal HandleValue As LongPtr, ByVal OptionCode As Long, ByVal OptionValue As Long) As Boolean
    SetOptionLong = (WinHttpSetOption(HandleValue, OptionCode, OptionValue, 4) <> 0)
End Function

Public Sub CopyByteRange(ByRef Destination() As Byte, ByVal DestinationOffset As Long, ByRef Source() As Byte, ByVal ByteCount As Long)
    If ByteCount <= 0 Then Exit Sub
    CopyMemory Destination(DestinationOffset), Source(0), ByteCount
End Sub
#Else
Public Function OpenSession(ByVal UserAgent As String) As Long
    OpenSession = WinHttpOpen(StrPtr(UserAgent), WinHttpAccessTypeDefaultProxy, 0, 0, 0)
End Function

Public Function Connect(ByVal SessionHandle As Long, ByVal ServerName As String, ByVal ServerPort As Long) As Long
    Connect = WinHttpConnect(SessionHandle, StrPtr(ServerName), ServerPort, 0)
End Function

Public Function OpenRequest(ByVal ConnectionHandle As Long, ByVal Method As String, ByVal ObjectName As String, ByVal Flags As Long) As Long
    OpenRequest = WinHttpOpenRequest(ConnectionHandle, StrPtr(Method), StrPtr(ObjectName), 0, 0, 0, Flags)
End Function

Public Function CloseHandle(ByVal HandleValue As Long) As Boolean
    If HandleValue = 0 Then
        CloseHandle = True
    Else
        CloseHandle = (WinHttpCloseHandle(HandleValue) <> 0)
    End If
End Function

Public Function SetTimeouts(ByVal HandleValue As Long, ByVal ResolveMilliseconds As Long, ByVal ConnectMilliseconds As Long, ByVal SendMilliseconds As Long, ByVal ReceiveMilliseconds As Long) As Boolean
    SetTimeouts = (WinHttpSetTimeouts(HandleValue, ResolveMilliseconds, ConnectMilliseconds, SendMilliseconds, ReceiveMilliseconds) <> 0)
End Function

Public Function AddRequestHeaders(ByVal HandleValue As Long, ByVal Headers As String) As Boolean
    If Len(Headers) = 0 Then
        AddRequestHeaders = True
    Else
        AddRequestHeaders = (WinHttpAddRequestHeaders(HandleValue, StrPtr(Headers), Len(Headers), WinHttpAddRequestHeader) <> 0)
    End If
End Function

Public Function SendBody(ByVal HandleValue As Long, ByVal Body As Variant) As Boolean
    Dim bytes() As Byte
    Dim lowerBound As Long
    Dim byteCount As Long

    If Not HttpEncoding.HasByteArray(Body) Then
        SendBody = (WinHttpSendRequest(HandleValue, 0, 0, 0, 0, 0, 0) <> 0)
        Exit Function
    End If
    bytes = Body
    lowerBound = LBound(bytes) ' xlflow:disable-line VBA227
    byteCount = UBound(bytes) - lowerBound + 1 ' xlflow:disable-line VBA227
    SendBody = (WinHttpSendRequest(HandleValue, 0, 0, VarPtr(bytes(lowerBound)), byteCount, byteCount, 0) <> 0) ' xlflow:disable-line VBA227
End Function

Public Function ReceiveResponse(ByVal HandleValue As Long) As Boolean
    ReceiveResponse = (WinHttpReceiveResponse(HandleValue, 0) <> 0)
End Function

Public Function QueryHeaderString(ByVal HandleValue As Long, ByVal InfoLevel As Long, ByRef Value As String) As Boolean
    Dim requiredBytes As Long
    Dim headerIndex As Long
    Dim buffer As String
    Dim errorCode As Long

    requiredBytes = 0
    headerIndex = 0
    If WinHttpQueryHeaders(HandleValue, InfoLevel, 0, 0, requiredBytes, headerIndex) <> 0 Then
        Value = ""
        QueryHeaderString = True
        Exit Function
    End If
    errorCode = GetLastError()
    If errorCode = WinHttpErrorHeaderNotFound Then
        Value = ""
        QueryHeaderString = True
        Exit Function
    End If
    If errorCode <> WinHttpErrorInsufficientBuffer Or requiredBytes <= 0 Or (requiredBytes Mod 2) <> 0 Then Exit Function

    buffer = String$(requiredBytes \ 2, vbNullChar)
    If WinHttpQueryHeaders(HandleValue, InfoLevel, 0, StrPtr(buffer), requiredBytes, headerIndex) = 0 Then Exit Function
    Value = RemoveTrailingNull(buffer)
    QueryHeaderString = True
End Function

Public Function QueryDataAvailable(ByVal HandleValue As Long, ByRef ByteCount As Long) As Boolean
    QueryDataAvailable = (WinHttpQueryDataAvailable(HandleValue, ByteCount) <> 0)
End Function

Public Function ReadData(ByVal HandleValue As Long, ByRef Buffer() As Byte, ByVal RequestedBytes As Long, ByRef ReadBytes As Long) As Boolean
    ReadData = (WinHttpReadData(HandleValue, Buffer(0), RequestedBytes, ReadBytes) <> 0)
End Function

Public Function QueryProtocolUsed(ByVal HandleValue As Long, ByRef ProtocolFlags As Long) As Boolean
    Dim bufferLength As Long

    ProtocolFlags = 0
    bufferLength = 4
    QueryProtocolUsed = (WinHttpQueryOption(HandleValue, WinHttpOptionHttpProtocolUsed, VarPtr(ProtocolFlags), bufferLength) <> 0)
End Function

Public Function SetOptionLong(ByVal HandleValue As Long, ByVal OptionCode As Long, ByVal OptionValue As Long) As Boolean
    SetOptionLong = (WinHttpSetOption(HandleValue, OptionCode, OptionValue, 4) <> 0)
End Function

Public Sub CopyByteRange(ByRef Destination() As Byte, ByVal DestinationOffset As Long, ByRef Source() As Byte, ByVal ByteCount As Long)
    If ByteCount <= 0 Then Exit Sub
    CopyMemory Destination(DestinationOffset), Source(0), ByteCount
End Sub
#End If

Public Function LastErrorCode() As Long
    LastErrorCode = Err.LastDllError
    If LastErrorCode = 0 Then LastErrorCode = GetLastError()
End Function

Public Function ProcessHandleCount(ByRef Count As Long) As Boolean
    Count = 0
    #If VBA7 Then
    ProcessHandleCount = (GetProcessHandleCount(GetCurrentProcess(), Count) <> 0)
    #Else
    ProcessHandleCount = (GetProcessHandleCount(GetCurrentProcess(), Count) <> 0)
    #End If
End Function

Private Function RemoveTrailingNull(ByVal Value As String) As String
    Dim nullIndex As Long

    nullIndex = InStr(1, Value, vbNullChar, vbBinaryCompare)
    If nullIndex > 0 Then
        RemoveTrailingNull = Left$(Value, nullIndex - 1)
    Else
        RemoveTrailingNull = Value
    End If
End Function
