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

Public Sub terminatingIndexedTickingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal callbackParams As UnmanagedCallbackWrapper, ByVal tickCount As Long)

    'this toggle makes sure TickerAPI is aware of any timers following a state change - it can then shut them down and lock out any bad behaviour (re-starts)
    Static timerChecked As Boolean 'should start False
    Dim data As Dictionary
    Debug.Print "DEBUG... ";
    Set data = Cache.loadObject("TickerApi.TimerIDs", New Dictionary)
    If Not timerChecked Then
        'Debug.Print callbackParams.timerID
        Debug.Print data.Count;
        Debug.Print "PRE-POKE"
        TickerAPI.Poke
        Debug.Print data.Count;
        Debug.Print "POST-POKE"
        timerChecked = True
    End If
    'Initialise dict {id:counter} with count of zero
    Static timerSet As New Dictionary 'should persist between callbacks but not over state change
    
    Debug.Print data.Count;
    If data.Count = 1 Then
        On Error Resume Next
        Debug.Print data.Keys(0);
        Debug.Print TypeName(data.Items(0)) & " ";
        Debug.Print data.Items(0).debugName;
        Debug.Print data.Items(0).timerID;
        On Error GoTo 0
    End If
    If Not timerSet.Exists(callbackParams.timerID) Then
        timerSet(callbackParams.timerID) = 0
    End If

    timerSet(callbackParams.timerID) = timerSet(callbackParams.timerID) + 1
    
    On Error GoTo 0
    Debug.Print Toolbox.Strings.Format("Ticking - {0} ({3}-id:{1})\t{2}", timerSet(callbackParams.timerID), callbackParams.timerID, time$, callbackParams.debugName)
    
    'Terminate timers which reach the max count
    If timerSet(callbackParams.timerID) >= 10 Then
        On Error Resume Next
        TickerAPI.KillTimerByID callbackParams.timerID          'stop timer
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

Public Sub passByRefProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal callbackParams As UnmanagedCallbackWrapper, ByVal tickCount As Long)
    Debug.Print "Callback called " & time
    On Error Resume Next
    Dim wrapper As ICallbackWrapper
    Set wrapper = callbackParams
    TickerAPI.KillTimersByFunction wrapper.Callback
    Debug.Print IIf(Cache.loadObject("TickerAPI.timerIDs", New Dictionary).Count = 0, "It's cleared", "Still hanging around:(")
    Debug.Print callbackParams.debugName, callbackParams.storedData 'check if still there
    On Error GoTo 0
End Sub
