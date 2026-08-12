VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SequencedHttpTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHttpTransport
Implements IHttpAttemptTransport

Private mOutcomes As New Collection
Private mExecuteCount As Long

Public Property Get ExecuteCount() As Long
    ExecuteCount = mExecuteCount
End Property

Public Sub EnqueueResponse(ByVal statusCode As Long, Optional ByVal retryAfter As String = "")
    Dim response As New HttpResponse
    Dim Headers As New HttpHeaders
    Dim item As New HttpBatchItem

    If Len(retryAfter) > 0 Then Headers.SetValue "Retry-After", retryAfter
    response.Initialize statusCode, "", Headers
    item.InitializeSuccess mOutcomes.Count + 1, response
    mOutcomes.Add item
End Sub

Public Sub EnqueueFailure(ByVal errorNumber As Long, ByVal Description As String)
    Dim item As New HttpBatchItem

    item.InitializeFailure mOutcomes.Count + 1, errorNumber, "SequencedHttpTransport.Execute", Description
    mOutcomes.Add item
End Sub

Private Function IHttpTransport_Execute(ByVal Request As HttpRequest) As HttpResponse
    Set IHttpTransport_Execute = NextOutcome(Request)
End Function

Private Function IHttpAttemptTransport_ExecuteAttempt(ByVal Request As HttpRequest, ByVal Options As HttpAttemptOptions) As HttpResponse
    If Not Options.CancellationToken Is Nothing Then
        If Options.CancellationToken.IsCancellationRequested Then Err.Raise HttpErrCancelled, "SequencedHttpTransport.Execute", "HTTP operation was cancelled."
    End If
    Set IHttpAttemptTransport_ExecuteAttempt = NextOutcome(Request)
End Function

Private Function NextOutcome(ByVal Request As HttpRequest) As HttpResponse
    Dim item As HttpBatchItem

    If Request Is Nothing Then Err.Raise HttpErrValidation, "SequencedHttpTransport.Execute", "Request cannot be Nothing."
    mExecuteCount = mExecuteCount + 1
    If mExecuteCount > mOutcomes.Count Then Err.Raise HttpErrProtocol, "SequencedHttpTransport.Execute", "No sequenced outcome remains."
    Set item = mOutcomes(mExecuteCount)
    If item.Status = HttpBatchSucceeded Then
        Set NextOutcome = item.Response.Clone()
    Else
        Err.Raise item.ErrorNumber, item.ErrorSource, item.ErrorDescription
    End If
End Function
