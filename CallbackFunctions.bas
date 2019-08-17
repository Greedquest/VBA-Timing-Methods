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

Public Sub terminatingIndexedTickingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByRef params As UnmanagedCallbackWrapper, ByVal tickCount As Long)

    Static timerChecked As Boolean 'should start False
    
'    If Not timerChecked Then
'        Debug.Assert True
'    End If
'    On Error Resume Next
'    Dim a As UnmanagedCallbackWrapper
'    Debug.Print "Dereferencing - ";
'    Set a = UnmanagedCallbackWrapper.FromPtr(params.timerID)
'    If Err.Number <> 0 Then
'        Debug.Print printf("Couldn't deref {0} - Err#{1}: {2}", params.timerID, Err.Number, Err.Description)
'    End If
'    On Error GoTo 0
   
    'this toggle makes sure TickerAPI is aware of any timers following a state change - it can then shut them down and lock out any bad behaviour (re-starts)
    If Not timerChecked Then
        Debug.Print UCase$("pre-poke")
        TickerAPI.Poke
        Debug.Print UCase$("post-poke")
        timerChecked = True
    End If
        
    Static timerSet As New Dictionary
    If Not timerSet.Exists(params.timerID) Then
        On Error Resume Next 'race contdition
        timerSet.item(params.timerID) = 0
        On Error GoTo 0
    End If
    timerSet(params.timerID) = timerSet(params.timerID) + 1
        
    'Debug.Print printf("Ticking - {0} (id:{1})", timerSet(timerID), timerID), time$
    
    Debug.Print printf("Ticking - {0} (id:{1})", timerSet(params.timerID), params.timerID), time$;
    If Not params Is Nothing Then
        Debug.Print " - "; params.storedData;
        params.storeData printf("Data Name: {0} Time:{1}", params.debugName, time$)
    End If
    Debug.Print 'for linefeed
        
    'Terminate by ID
    If timerSet(params.timerID) >= 10 Then
        On Error Resume Next
        TickerAPI.KillTimerByID params.timerID          'stop timer
        On Error GoTo 0
        timerSet.Remove params.timerID
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

Public Sub passByRefProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByRef params As UnmanagedCallbackWrapper, ByVal tickCount As Long)
    Debug.Print "Callback called " & time
    On Error Resume Next
    TickerAPI.KillTimerByID timerID
    On Error GoTo 0
End Sub
