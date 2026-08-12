Attribute VB_Name = "HttpTiming"
Option Explicit

Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef value As HttpSystemTime)
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef value As HttpSystemTime)
#End If

Private Type HttpSystemTime
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type

Public Function CounterValue() As Currency
    If QueryPerformanceCounter(CounterValue) = 0 Then
        Err.Raise HttpErrIo, "HttpTiming.CounterValue", "High-resolution timer is unavailable."
    End If
End Function

Public Function ElapsedMilliseconds(ByVal started As Currency, ByVal finished As Currency) As Double
    Static frequency As Currency

    If frequency = 0 Then
        If QueryPerformanceFrequency(frequency) = 0 Then
            Err.Raise HttpErrIo, "HttpTiming.ElapsedMilliseconds", "High-resolution timer is unavailable."
        End If
    End If
    ElapsedMilliseconds = (CDbl(finished) - CDbl(started)) * 1000# / CDbl(frequency)
End Function

Public Sub Pause(ByVal milliseconds As Long)
    If milliseconds < 0 Then HttpErrors.RaiseValidation "HttpTiming.Pause", "Pause duration cannot be negative."
    If milliseconds > 0 Then Sleep milliseconds
End Sub

Public Function UtcNow() As Date
    Dim value As HttpSystemTime

    GetSystemTime value
    UtcNow = DateSerial(value.Year, value.Month, value.Day) + TimeSerial(value.Hour, value.Minute, value.Second) + CDbl(value.Milliseconds) / 86400000#
End Function
