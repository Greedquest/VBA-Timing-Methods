Attribute VB_Name = "CallbackFunctions"
Option Explicit
'@Folder("Tests.Callbacks")
'@IgnoreModule ParameterNotUsed, ProcedureCanBeWrittenAsFunction

'Public Type TCallbackSettings
'    sourceNames As New Dictionary
'    defaultMaxTicks As Long
'End Type
'
'Public callbackSettings As TCallbackSettings

Public Sub SafeCallbackProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    Debug.Print "Callback called " & time
    On Error Resume Next
    TickerAPI.KillTimerByID timerID
    On Error GoTo 0
End Sub

Public Sub QuietTerminatingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    On Error Resume Next
    TickerAPI.KillTimerByID timerID
    On Error GoTo 0
End Sub

Public Sub QuietNoOpCallback(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
End Sub

Public Sub RawSelfKillingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    Debug.Print "Tick"
    WinAPI.KillTimer windowHandle, timerID
End Sub

'@Description("Ticks with automatic termination")
Public Sub SafeTickingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
Attribute SafeTickingProc.VB_Description = "Ticks with automatic termination"
    Static i As Long
    Debug.Print "Ticking "; i
    i = i + 1
    If i > 10 Then
        On Error Resume Next
        TickerAPI.KillTimerByID timerID          'stop timer
        On Error GoTo 0
    End If
End Sub

Public Sub terminatingIndexedTickingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)

    Static timerSet As New Dictionary
    If Not timerSet.Exists(timerID) Then timerSet.Add timerID, 0
    timerSet(timerID) = timerSet(timerID) + 1
        
    Debug.Print printf("Ticking - {0} (id:{1})", timerSet(timerID), timerID), time$
    If timerSet(timerID) > 10 Then
        On Error Resume Next
        TickerAPI.KillTimerByID timerID          'stop timer
        On Error GoTo 0
    End If
    
End Sub

Public Sub RecursiveProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
'NOTE no recursion anymore
    Static i As Long
    i = i + 1
    Debug.Print i; "Callback called " & time; timerID
    If i < 3 Then TickerAPI.StartUnmanagedTimer AddressOf RecursiveProc, , True, 1000
    Debug.Print i
    i = i - 1
    'createTimer.TickerIsRunning = i = 1
End Sub

