VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpCookieJar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MaximumCookies As Long = 300
Private Const MaximumSerializedCookieBytes As Long = 4096

Private mCookies As Collection
Private mClock As IHttpCookieClock

Public Property Get Count() As Long
    Count = mCookies.Count
End Property

Public Property Set Clock(ByVal value As IHttpCookieClock)
    If value Is Nothing Then
        Set mClock = Nothing
    Else
        Set mClock = value
    End If
End Property

Public Sub Clear()
    Set mCookies = New Collection
End Sub

Public Function GetCookieHeader(ByVal Url As String) As String
    Dim target As HttpCookieUrl
    Dim nowUtc As Date
    Dim matches As Collection
    Dim item As HttpCookieRecord
    Dim output As String

    Set target = HttpCookieParsing.ParseUrl(Url)
    nowUtc = CurrentUtc()
    RemoveExpired nowUtc
    Set matches = MatchingCookies(target, nowUtc)
    For Each item In matches
        If Len(output) > 0 Then output = output & "; "
        output = output & item.Name & "=" & item.Value
    Next item
    GetCookieHeader = output
End Function

Public Sub StoreResponseCookies(ByVal Url As String, ByVal Headers As HttpHeaders)
    Dim target As HttpCookieUrl
    Dim nowUtc As Date
    Dim index As Long
    Dim item As HttpCookieRecord

    If Headers Is Nothing Then Exit Sub
    Set target = HttpCookieParsing.ParseUrl(Url)
    nowUtc = CurrentUtc()
    For index = 1 To Headers.Count
        If StrComp(Headers.NameAt(index), "Set-Cookie", vbTextCompare) = 0 Then
            Set item = HttpCookieParsing.TryParseSetCookie(Headers.ValueAt(index), target, nowUtc)
            If Not item Is Nothing Then StoreOrDelete item, nowUtc
        End If
    Next index
End Sub

Private Sub StoreOrDelete(ByVal item As HttpCookieRecord, ByVal nowUtc As Date)
    Dim index As Long

    If item.SerializedLength > MaximumSerializedCookieBytes Then
        HttpErrors.RaiseValidation "HttpCookieJar.StoreResponseCookies", "Cookie exceeds the 4096-byte limit."
    End If
    index = FindCookie(item.Name, item.Domain, item.Path)
    If item.IsExpired(nowUtc) Then
        If index > 0 Then mCookies.Remove index
        Exit Sub
    End If
    If index > 0 Then mCookies.Remove index
    If mCookies.Count >= MaximumCookies Then
        HttpErrors.RaiseValidation "HttpCookieJar.StoreResponseCookies", "Cookie jar capacity is 300 cookies."
    End If
    mCookies.Add item
End Sub

Private Function MatchingCookies(ByVal target As HttpCookieUrl, ByVal nowUtc As Date) As Collection
    Dim output As New Collection
    Dim item As HttpCookieRecord
    Dim index As Long

    For index = 1 To mCookies.Count
        Set item = mCookies(index)
        If item.Matches(target, nowUtc) Then InsertByPath output, item
    Next index
    Set MatchingCookies = output
End Function

Private Sub InsertByPath(ByVal items As Collection, ByVal item As HttpCookieRecord)
    Dim index As Long
    Dim candidate As HttpCookieRecord

    If items Is Nothing Then Set items = New Collection
    If item Is Nothing Then Exit Sub
    For index = 1 To items.Count
        Set candidate = items(index)
        If Len(item.Path) > Len(candidate.Path) Then
            items.Add item, , index
            Exit Sub
        End If
    Next index
    items.Add item
End Sub

Private Sub RemoveExpired(ByVal nowUtc As Date)
    Dim index As Long
    Dim item As HttpCookieRecord

    index = mCookies.Count
    Do While index >= 1
        Set item = mCookies(index)
        If item.IsExpired(nowUtc) Then mCookies.Remove index
        index = index - 1
    Loop
End Sub

Private Function FindCookie(ByVal Name As String, ByVal Domain As String, ByVal Path As String) As Long
    Dim index As Long
    Dim item As HttpCookieRecord

    For index = 1 To mCookies.Count
        Set item = mCookies(index)
        If StrComp(item.Name, Name, vbTextCompare) = 0 And _
            StrComp(item.Domain, Domain, vbTextCompare) = 0 And _
            item.Path = Path Then
            FindCookie = index
            Exit Function
        End If
    Next index
End Function

Private Function CurrentUtc() As Date
    If mClock Is Nothing Then
        CurrentUtc = HttpTiming.UtcNow()
    Else
        CurrentUtc = mClock.UtcNow()
    End If
End Function

Private Sub Class_Initialize()
    Set mCookies = New Collection
End Sub
