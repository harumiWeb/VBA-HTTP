VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpMultipartPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIsFile As Boolean
Private mName As String
Private mValue As String
Private mFilePath As String
Private mFileName As String
Private mContentType As String

Public Property Get IsFile() As Boolean
    IsFile = mIsFile
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get Value() As String
    If mIsFile Then HttpErrors.RaiseValidation "HttpMultipartPart.Value", "File parts do not expose a field value."
    Value = mValue
End Property

Public Property Get FilePath() As String
    If Not mIsFile Then HttpErrors.RaiseValidation "HttpMultipartPart.FilePath", "Field parts do not expose a file path."
    FilePath = mFilePath
End Property

Public Property Get FileName() As String
    If Not mIsFile Then HttpErrors.RaiseValidation "HttpMultipartPart.FileName", "Field parts do not expose a file name."
    FileName = mFileName
End Property

Public Property Get ContentType() As String
    If Not mIsFile Then HttpErrors.RaiseValidation "HttpMultipartPart.ContentType", "Field parts do not expose a content type."
    ContentType = mContentType
End Property

Friend Sub InitializeField(ByVal Name As String, ByVal Value As String)
    mIsFile = False
    mName = Name
    mValue = Value
    mFilePath = ""
    mFileName = ""
    mContentType = ""
End Sub

Friend Sub InitializeFile(ByVal Name As String, ByVal FilePath As String, ByVal FileName As String, ByVal ContentType As String)
    mIsFile = True
    mName = Name
    mValue = ""
    mFilePath = FilePath
    mFileName = FileName
    mContentType = ContentType
End Sub

Friend Function Clone() As HttpMultipartPart
    Dim copy As New HttpMultipartPart

    If mIsFile Then
        copy.InitializeFile mName, mFilePath, mFileName, mContentType
    Else
        copy.InitializeField mName, mValue
    End If
    Set Clone = copy
End Function
