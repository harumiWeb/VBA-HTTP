VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockHttpTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHttpTransport

Private mResponse As HttpResponse
Private mLastRequest As HttpRequest
Private mExecuteCount As Long
Private mErrorNumber As Long
Private mErrorDescription As String

Public Property Get LastRequest() As HttpRequest
    Set LastRequest = mLastRequest
End Property

Public Property Get ExecuteCount() As Long
    ExecuteCount = mExecuteCount
End Property

Public Sub SetResponse(ByVal value As HttpResponse)
    If value Is Nothing Then
        HttpErrors.RaiseValidation "MockHttpTransport.SetResponse", "Response cannot be Nothing."
    End If
    Set mResponse = value.Clone()
    mErrorNumber = 0
    mErrorDescription = ""
End Sub

Public Sub SetFailure(ByVal errorNumber As Long, ByVal Description As String)
    mErrorNumber = errorNumber
    mErrorDescription = Description
End Sub

Private Function IHttpTransport_Execute(ByVal Request As HttpRequest) As HttpResponse
    mExecuteCount = mExecuteCount + 1
    Set mLastRequest = Request.Clone()
    If mErrorNumber <> 0 Then
        Err.Raise mErrorNumber, "MockHttpTransport.Execute", mErrorDescription
    End If
    If mResponse Is Nothing Then
        Err.Raise HttpErrProtocol, "MockHttpTransport.Execute", "Mock response was not configured."
    End If
    Set IHttpTransport_Execute = mResponse.Clone()
End Function
