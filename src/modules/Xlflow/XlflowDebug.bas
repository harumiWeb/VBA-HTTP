Attribute VB_Name = "XlflowDebug"
Option Explicit

' XlflowDebug mirrors workbook-side debug output to the terminal during xlflow runs.
' Use XlflowDebug.Log instead of raw Debug.Print when terminal-visible logs are desired.
Private Const xlflowDebugPipeName As String = "__XLFLOW_DEBUG_PIPE__"

#If VBA7 Then
	Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
	Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
	Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
	Private Const xlflowInvalidHandleValue As LongPtr = -1
#Else
	Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
	Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
	Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
	Private Const xlflowInvalidHandleValue As Long = -1
#End If

Private Const xlflowGenericWrite As Long = &H40000000
Private Const xlflowOpenExisting As Long = 3

''' Writes workbook-side debug output to Debug.Print and xlflow run/test output when available.
'''
''' Args:
'''     Parts: Values to stringify and join with spaces.
Public Sub Log(ParamArray Parts() As Variant)
	Dim index As Long
	Dim lowerBound As Long
	Dim upperBound As Long
	Dim message As String
	Dim errorNumber As Long
	Dim errorSource As String
	Dim errorDescription As String

	' ParamArray is allocated when arguments exist; error 9 below handles an empty argument list.
	On Error GoTo EmptyParts
	lowerBound = LBound(Parts) ' xlflow:disable-line VBA227
	upperBound = UBound(Parts) ' xlflow:disable-line VBA227
	On Error GoTo 0

	For index = lowerBound To upperBound ' xlflow:disable-line VBA227
		If index > lowerBound Then
			message = message & " "
		End If
		message = message & StringifyValue(Parts(index)) ' xlflow:disable-line VBA227
	Next index

GoTo PrintMessage

EmptyParts:
	errorNumber = Err.Number
	errorSource = Err.Source
	errorDescription = Err.Description
	Err.Clear
	On Error GoTo 0
	If errorNumber <> 9 Then
		Err.Raise errorNumber, errorSource, errorDescription
	End If

PrintMessage:
	If Len(message) = 0 Then
		Debug.Print
	Else
		Debug.Print message
	End If

	EmitDebugEvent message
End Sub

Private Sub EmitDebugEvent(ByVal Message As String)
	Dim pipeName As String
	Dim payload As String

	pipeName = ResolveDebugPipeName()
	If Len(pipeName) = 0 Then
		Exit Sub
	End If

	payload = "{" & _
		JsonProperty("event", "debug_log") & "," & _
		JsonProperty("message", Message) & "," & _
		JsonProperty("runtime_mode", XlflowRuntime.ModeName()) & "," & _
		JsonProperty("source", "XlflowDebug.Log") & "}"

	Call SendPipeText(pipeName, payload & vbLf)
End Sub

Private Function StringifyValue(ByVal Value As Variant) As String
	If IsObject(Value) Then
		On Error GoTo ObjectFallback
		StringifyValue = "[Object " & TypeName(Value) & "]"
		On Error GoTo 0
		Exit Function
ObjectFallback:
		Err.Clear
		On Error GoTo 0
		StringifyValue = "[Object]"
		Exit Function
	End If

	If IsArray(Value) Then
		StringifyValue = "[Array]"
		Exit Function
	End If

	If IsEmpty(Value) Then
		StringifyValue = "[Empty]"
		Exit Function
	End If

	If IsNull(Value) Then
		StringifyValue = "[Null]"
		Exit Function
	End If

	Select Case VarType(Value)
		Case vbBoolean
			If CBool(Value) Then
				StringifyValue = "True"
			Else
				StringifyValue = "False"
			End If
		Case vbError
			StringifyValue = "[Error " & CStr(CLng(Value)) & "]"
		Case Else
			On Error GoTo UnsupportedValue
			StringifyValue = CStr(Value)
			On Error GoTo 0
			Exit Function
UnsupportedValue:
			Err.Clear
			On Error GoTo 0
			StringifyValue = "[Unsupported Variant]"
	End Select
End Function

Private Function ResolveDebugPipeName() As String
	ResolveDebugPipeName = ReadOptionalDefinedNameValue(xlflowDebugPipeName)
End Function

Private Function ReadOptionalDefinedNameValue(ByVal Name As String) As String
	On Error GoTo Missing
	ReadOptionalDefinedNameValue = DecodeWorkbookDefinedName(ThisWorkbook.Names(Name).RefersTo)
	Exit Function

Missing:
	ReadOptionalDefinedNameValue = ""
End Function

Private Function DecodeWorkbookDefinedName(ByVal RefersTo As String) As String
	If Len(RefersTo) = 0 Then
		DecodeWorkbookDefinedName = ""
		Exit Function
	End If
	If Left$(RefersTo, 1) = "=" Then
		RefersTo = Mid$(RefersTo, 2)
	End If
	If Len(RefersTo) >= 2 Then
		If Left$(RefersTo, 1) = Chr$(34) And Right$(RefersTo, 1) = Chr$(34) Then
			RefersTo = Mid$(RefersTo, 2, Len(RefersTo) - 2)
		End If
	End If
	DecodeWorkbookDefinedName = Replace$(RefersTo, Chr$(34) & Chr$(34), Chr$(34))
End Function

Private Function JsonEscape(ByVal Value As String) As String
	Value = Replace$(Value, "\", "\\")
	Value = Replace$(Value, Chr$(34), Chr$(92) & Chr$(34))
	Value = Replace$(Value, vbCrLf, "\n")
	Value = Replace$(Value, vbCr, "\n")
	Value = Replace$(Value, vbLf, "\n")
	Value = Replace$(Value, vbTab, "\t")
	JsonEscape = Value
End Function

Private Function JsonProperty(ByVal Name As String, ByVal Value As String) As String
	JsonProperty = Chr$(34) & JsonEscape(Name) & Chr$(34) & ":" & Chr$(34) & JsonEscape(Value) & Chr$(34)
End Function

Private Function SendPipeText(ByVal PipeName As String, ByVal Payload As String) As Boolean
	Dim bytesWritten As Long
	Dim writeSucceeded As Long
#If VBA7 Then
	Dim pipeHandle As LongPtr
#Else
	Dim pipeHandle As Long
#End If

	SendPipeText = False
	pipeHandle = CreateFileW(StrPtr(PipeName), xlflowGenericWrite, 0, 0, xlflowOpenExisting, 0, 0)
	If pipeHandle = xlflowInvalidHandleValue Then
		Exit Function
	End If

	On Error GoTo Cleanup
	writeSucceeded = WriteFile(pipeHandle, StrPtr(Payload), Len(Payload) * 2, bytesWritten, 0)
	SendPipeText = writeSucceeded <> 0 And bytesWritten = Len(Payload) * 2

Cleanup:
	If pipeHandle <> xlflowInvalidHandleValue Then
		Call CloseHandle(pipeHandle)
	End If
	On Error GoTo 0
End Function
