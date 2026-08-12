VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpBatchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mItems As New Collection

Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Function ItemAt(ByVal index As Long) As HttpBatchItem
    If index < 1 Or index > mItems.Count Then HttpErrors.RaiseValidation "HttpBatchResult.ItemAt", "Batch item index is out of range."
    Set ItemAt = mItems(index)
End Function

Public Property Get SuccessCount() As Long
    SuccessCount = CountStatus(HttpBatchSucceeded)
End Property

Public Property Get FailureCount() As Long
    FailureCount = CountStatus(HttpBatchFailed)
End Property

Public Property Get CancelledCount() As Long
    CancelledCount = CountStatus(HttpBatchCancelled)
End Property

Friend Sub Add(ByVal item As HttpBatchItem)
    If item Is Nothing Then HttpErrors.RaiseValidation "HttpBatchResult.Add", "Batch item cannot be Nothing."
    If item.Index <> mItems.Count + 1 Then HttpErrors.RaiseValidation "HttpBatchResult.Add", "Batch items must be added in input order."
    mItems.Add item
End Sub

Private Function CountStatus(ByVal expected As HttpBatchItemStatus) As Long
    Dim item As HttpBatchItem

    For Each item In mItems
        If item.Status = expected Then CountStatus = CountStatus + 1
    Next item
End Function
