VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpCookieUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mScheme As String
Private mHost As String
Private mPath As String

Public Property Get Scheme() As String
    Scheme = mScheme
End Property

Public Property Let Scheme(ByVal value As String)
    mScheme = LCase$(value)
End Property

Public Property Get Host() As String
    Host = mHost
End Property

Public Property Let Host(ByVal value As String)
    mHost = LCase$(value)
End Property

Public Property Get Path() As String
    Path = mPath
End Property

Public Property Let Path(ByVal value As String)
    mPath = value
End Property

Public Property Get IsHttps() As Boolean
    IsHttps = (mScheme = "https")
End Property
