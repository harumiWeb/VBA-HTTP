Attribute VB_Name = "HttpConstants"
Option Explicit

Public Enum HttpBodyKind
    HttpBodyEmpty = 0
    HttpBodyText = 1
    HttpBodyBytes = 2
End Enum

Public Enum HttpErrorCategory
    HttpErrorNone = 0
    HttpErrorValidation = 1
    HttpErrorInvalidUrl = 2
    HttpErrorDns = 3
    HttpErrorConnection = 4
    HttpErrorTls = 5
    HttpErrorTimeout = 6
    HttpErrorCancelled = 7
    HttpErrorProtocol = 8
    HttpErrorIo = 9
    HttpErrorStatus = 10
End Enum

Public Enum HttpBatchItemStatus
    HttpBatchSucceeded = 1
    HttpBatchFailed = 2
    HttpBatchCancelled = 3
End Enum

Public Enum HttpProtocolFlag
    HttpProtocolHttp11 = 0
    HttpProtocolHttp2 = 1
    HttpProtocolHttp3 = 2
End Enum

Public Enum HttpProtocolMode
    HttpProtocolAllowFallback = 0
    HttpProtocolRequired = 1
End Enum

Public Enum HttpDecompressionFlag
    HttpDecompressionGzip = 1
    HttpDecompressionDeflate = 2
    HttpDecompressionAll = 3
End Enum

Public Enum HttpDecompressionMode
    HttpDecompressionAllowFallback = 0
    HttpDecompressionRequired = 1
End Enum

Public Enum HttpProxyMode
    HttpProxyDefault = 0
    HttpProxyNoProxy = 1
    HttpProxyManual = 2
End Enum
