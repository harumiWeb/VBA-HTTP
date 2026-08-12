Attribute VB_Name = "XlflowRuntime"
Option Explicit

' XlflowRuntime exposes the execution mode that xlflow injected before user VBA started.
' Use these helpers when workbook code must branch between interactive and unattended flows.
Private Const xlflowInteractive As Long = 0
Private Const xlflowHeadless As Long = 1
Private Const xlflowCI As Long = 2
Private Const xlflowAgent As Long = 3
Private Const xlflowTest As Long = 4

''' Returns the current xlflow runtime mode as a stable numeric value.
'''
''' Returns:
'''     One of the internal xlflow mode constants.
Public Function Mode() As Long
	Select Case ModeName()
		Case "headless"
			Mode = xlflowHeadless
		Case "ci"
			Mode = xlflowCI
		Case "agent"
			Mode = xlflowAgent
		Case "test"
			Mode = xlflowTest
		Case Else
			Mode = xlflowInteractive
	End Select
End Function

''' Returns the normalized runtime mode name injected by xlflow.
'''
''' Returns:
'''     interactive, headless, ci, agent, or test.
Public Function ModeName() As String
	Dim raw As String
	raw = ReadWorkbookModeName()
	If Len(raw) = 0 Then
		raw = Environ$("XLFLOW_MODE")
	End If
	raw = LCase$(Trim$(raw))

	Select Case raw
		Case "headless", "ci", "agent", "test"
			ModeName = raw
		Case Else
			ModeName = "interactive"
	End Select
End Function

''' Indicates whether the workbook is running in normal human-driven Excel usage.
'''
''' Returns:
'''     True when the runtime mode is interactive.
Public Function IsInteractive() As Boolean
	IsInteractive = (Mode() = xlflowInteractive)
End Function

''' Indicates whether the workbook is running without direct human interaction.
'''
''' Returns:
'''     True for headless, CI, agent, and test modes.
Public Function IsHeadless() As Boolean
	Select Case Mode()
		Case xlflowHeadless, xlflowCI, xlflowAgent, xlflowTest
			IsHeadless = True
		Case Else
			IsHeadless = False
	End Select
End Function

''' Indicates whether the workbook is running in CI mode.
'''
''' Returns:
'''     True when the runtime mode is ci.
Public Function IsCI() As Boolean
	IsCI = (Mode() = xlflowCI)
End Function

''' Indicates whether the workbook is running in agent mode.
'''
''' Returns:
'''     True when the runtime mode is agent.
Public Function IsAgent() As Boolean
	IsAgent = (Mode() = xlflowAgent)
End Function

''' Indicates whether the workbook is running under xlflow test.
'''
''' Returns:
'''     True when the runtime mode is test.
Public Function IsTest() As Boolean
	IsTest = (Mode() = xlflowTest)
End Function

Private Function ReadWorkbookModeName() As String
	On Error GoTo Missing
	ReadWorkbookModeName = DecodeWorkbookDefinedName(ThisWorkbook.Names("__XLFLOW_MODE__").RefersTo)
	Exit Function

Missing:
	ReadWorkbookModeName = ""
End Function

Private Function DecodeWorkbookDefinedName(ByVal refersTo As String) As String
	If Len(refersTo) = 0 Then
		DecodeWorkbookDefinedName = ""
		Exit Function
	End If
	If Left$(refersTo, 1) = "=" Then
		refersTo = Mid$(refersTo, 2)
	End If
	If Len(refersTo) >= 2 Then
		If Left$(refersTo, 1) = Chr$(34) And Right$(refersTo, 1) = Chr$(34) Then
			refersTo = Mid$(refersTo, 2, Len(refersTo) - 2)
		End If
	End If
	DecodeWorkbookDefinedName = Replace$(refersTo, Chr$(34) & Chr$(34), Chr$(34))
End Function
