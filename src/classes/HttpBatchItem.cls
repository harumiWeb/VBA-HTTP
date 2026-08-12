VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpBatchItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIndex As Long
Private mStatus As HttpBatchItemStatus
Private mResponse As HttpResponse
Private mErrorNumber As Long
Private mErrorCategory As HttpErrorCategory
Private mErrorSource As String
Private mErrorDescription As String

Public Property Get Index() As Long
    Index = mIndex
End Property

Public Property Get Status() As HttpBatchItemStatus
    Status = mStatus
End Property

Public Property Get Response() As HttpResponse
    Set Response = mResponse
End Property

Public Property Get ErrorNumber() As Long
    ErrorNumber = mErrorNumber
End Property

Public Property Get ErrorCategory() As HttpErrorCategory
    ErrorCategory = mErrorCategory
End Property

Public Property Get ErrorSource() As String
    ErrorSource = mErrorSource
End Property

Public Property Get ErrorDescription() As String
    ErrorDescription = mErrorDescription
End Property

Friend Sub InitializeSuccess(ByVal itemIndex As Long, ByVal value As HttpResponse)
    If value Is Nothing Then HttpErrors.RaiseValidation "HttpBatchItem.InitializeSuccess", "Successful batch item requires a response."
    mIndex = itemIndex
    mStatus = HttpBatchSucceeded
    Set mResponse = value
End Sub

Friend Sub InitializeFailure(ByVal itemIndex As Long, ByVal number As Long, ByVal Source As String, ByVal Description As String)
    mIndex = itemIndex
    mStatus = HttpBatchFailed
    mErrorNumber = number
    mErrorCategory = HttpErrors.CategoryFromNumber(number)
    mErrorSource = Source
    mErrorDescription = Description
End Sub

Friend Sub InitializeCancelled(ByVal itemIndex As Long)
    mIndex = itemIndex
    mStatus = HttpBatchCancelled
    mErrorNumber = HttpErrCancelled
    mErrorCategory = HttpErrorCancelled
    mErrorSource = "HttpClient.ExecuteMany"
    mErrorDescription = "HTTP batch item was cancelled."
End Sub
