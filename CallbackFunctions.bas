Attribute VB_Name = "CallbackFunctions"
Option Explicit
'@Folder("Tests.Experiments.Callbacks")
'@IgnoreModule ParameterNotUsed, ProcedureCanBeWrittenAsFunction

'Public Type TCallbackSettings
'    sourceNames As New Dictionary
'    defaultMaxTicks As Long
'End Type
'
'Public callbackSettings As TCallbackSettings

Public Sub SafeCallbackProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal callbackParams As UnmanagedCallbackWrapper, ByVal tickCount As Long)
    
    Dim expectedData As String
    On Error Resume Next
    expectedData = CStr(callbackParams.userData)
    On Error GoTo 0
    
    Debug.Print "Callback called " & time & " Data: '" & expectedData & "'"
    
    On Error Resume Next
    TickerAPI.KillTimerByID callbackParams.timerID
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
    WinAPI.killTimer windowHandle, timerID
End Sub

'@Description("Ticks with automatic termination")
Public Sub SafeTickingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    Static i As Long
    Debug.Print "Ticking "; i
    i = i + 1
    If i > 10 Then
        On Error Resume Next
        TickerAPI.KillTimerByID timerID          'stop timer
        On Error GoTo 0
    End If
End Sub

Public Sub terminatingIndexedTickingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal callbackParams As UnmanagedCallbackWrapper, ByVal tickCount As Long)

    'Initialise dict {id:counter} with count of zero
    Static timerSet As New Dictionary            'should persist between callbacks but not over state change
   
    'Increment counter stuff
    If Not timerSet.Exists(callbackParams.timerID) Then
        timerSet(callbackParams.timerID) = 0
    End If
    timerSet(callbackParams.timerID) = timerSet(callbackParams.timerID) + 1
    
    'Get & Log info
    Dim expectedData As String
    On Error Resume Next
    expectedData = CStr(callbackParams.userData) 'catch error in case of bad data
    On Error GoTo 0
    
    Debug.Print Toolbox.Strings.Format("Ticking - {0} (id:{1})\tData:'{3}'\t{2}", timerSet(callbackParams.timerID), callbackParams.timerID, time$, expectedData)
    
    'Terminate timers which reach the max count
    If timerSet(callbackParams.timerID) >= 10 Then
        On Error Resume Next
        TickerAPI.KillTimerByID callbackParams.timerID 'stop timer
        On Error GoTo 0
        timerSet.Remove callbackParams.timerID
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

Public Sub RawSafeTickingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    Static i As Long
    i = i + 1
    Debug.Print i; "Tick"
    
    If i >= 10 Then
        Debug.Print "Terminating"
        WinAPI.killTimer windowHandle, timerID
    End If
End Sub
