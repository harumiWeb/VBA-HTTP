Attribute VB_Name = "HttpErrors"
Option Explicit

Public Const HttpErrorBase As Long = vbObjectError + 21000
Public Const HttpErrValidation As Long = HttpErrorBase + 1
Public Const HttpErrInvalidUrl As Long = HttpErrorBase + 2
Public Const HttpErrDns As Long = HttpErrorBase + 3
Public Const HttpErrConnection As Long = HttpErrorBase + 4
Public Const HttpErrTls As Long = HttpErrorBase + 5
Public Const HttpErrTimeout As Long = HttpErrorBase + 6
Public Const HttpErrCancelled As Long = HttpErrorBase + 7
Public Const HttpErrProtocol As Long = HttpErrorBase + 8
Public Const HttpErrIo As Long = HttpErrorBase + 9
Public Const HttpErrStatus As Long = HttpErrorBase + 10

Public Sub RaiseValidation(ByVal Source As String, ByVal Description As String)
    Err.Raise HttpErrValidation, Source, Description
End Sub

Public Sub RaiseInvalidUrl(ByVal Source As String, ByVal Description As String)
    Err.Raise HttpErrInvalidUrl, Source, Description
End Sub

Public Sub RaiseStatus(ByVal Source As String, ByVal statusCode As Long)
    Err.Raise HttpErrStatus, Source, "HTTP request returned status " & CStr(statusCode) & "."
End Sub

Public Function CategoryFromNumber(ByVal errorNumber As Long) As HttpErrorCategory
    Select Case errorNumber
    Case HttpErrValidation
        CategoryFromNumber = HttpErrorValidation
    Case HttpErrInvalidUrl
        CategoryFromNumber = HttpErrorInvalidUrl
    Case HttpErrDns
        CategoryFromNumber = HttpErrorDns
    Case HttpErrConnection
        CategoryFromNumber = HttpErrorConnection
    Case HttpErrTls
        CategoryFromNumber = HttpErrorTls
    Case HttpErrTimeout
        CategoryFromNumber = HttpErrorTimeout
    Case HttpErrCancelled
        CategoryFromNumber = HttpErrorCancelled
    Case HttpErrProtocol
        CategoryFromNumber = HttpErrorProtocol
    Case HttpErrIo
        CategoryFromNumber = HttpErrorIo
    Case HttpErrStatus
        CategoryFromNumber = HttpErrorStatus
    Case Else
        CategoryFromNumber = HttpErrorNone
    End Select
End Function

Public Function NumberFromCategory(ByVal category As HttpErrorCategory) As Long
    Select Case category
    Case HttpErrorNone
        NumberFromCategory = 0
    Case HttpErrorValidation
        NumberFromCategory = HttpErrValidation
    Case HttpErrorInvalidUrl
        NumberFromCategory = HttpErrInvalidUrl
    Case HttpErrorDns
        NumberFromCategory = HttpErrDns
    Case HttpErrorConnection
        NumberFromCategory = HttpErrConnection
    Case HttpErrorTls
        NumberFromCategory = HttpErrTls
    Case HttpErrorTimeout
        NumberFromCategory = HttpErrTimeout
    Case HttpErrorCancelled
        NumberFromCategory = HttpErrCancelled
    Case HttpErrorProtocol
        NumberFromCategory = HttpErrProtocol
    Case HttpErrorIo
        NumberFromCategory = HttpErrIo
    Case HttpErrorStatus
        NumberFromCategory = HttpErrStatus
    Case Else
        RaiseValidation "HttpErrors.NumberFromCategory", "Unknown HTTP error category."
    End Select
End Function

Public Function CategoryName(ByVal category As HttpErrorCategory) As String
    Select Case category
    Case HttpErrorNone
        CategoryName = "none"
    Case HttpErrorValidation
        CategoryName = "validation"
    Case HttpErrorInvalidUrl
        CategoryName = "invalid_url"
    Case HttpErrorDns
        CategoryName = "dns"
    Case HttpErrorConnection
        CategoryName = "connection"
    Case HttpErrorTls
        CategoryName = "tls"
    Case HttpErrorTimeout
        CategoryName = "timeout"
    Case HttpErrorCancelled
        CategoryName = "cancelled"
    Case HttpErrorProtocol
        CategoryName = "protocol"
    Case HttpErrorIo
        CategoryName = "io"
    Case HttpErrorStatus
        CategoryName = "status"
    Case Else
        CategoryName = "unknown"
    End Select
End Function

Public Sub RaiseTransport(ByVal category As HttpErrorCategory, ByVal Source As String, ByVal Description As String)
    Select Case category
    Case HttpErrorInvalidUrl, HttpErrorDns, HttpErrorConnection, HttpErrorTls, HttpErrorTimeout, HttpErrorCancelled, HttpErrorProtocol, HttpErrorIo
        Err.Raise NumberFromCategory(category), Source, Description
    Case Else
        RaiseValidation "HttpErrors.RaiseTransport", "Category is not a transport failure."
    End Select
End Sub
