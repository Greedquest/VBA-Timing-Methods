Attribute VB_Name = "MessageWindowProcs"
'@Folder("FirstLevelAPI.Utils")
Option Explicit
Option Private Module

Private activated As Boolean

Public Sub Activate()
    'TODO a cleaner way would be to ["Generate completely unique message"](https://stackoverflow.com/q/57625385/6609896) and pass that to the subclass proc so it can keep a static variable
    activated = True 'auto resets to False on state loss
End Sub

Public Function ManagedTimerMessageWindowSubclassProc(ByVal hWnd As LongPtr, ByVal uMsg As WindowsMessage, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr

    'Debug.Print "MSG #"; uMsg, activated
    Select Case uMsg
    
            'NOTE this will never receive timer messages where TIMERPROC is specified,
        Case WindowsMessage.WM_TIMER             'wParam = timerID , lParam = "timerProc" (will be Null if it reaches here)
            If activated Then
            'NOTE TickerAPI is pre-declared so checking it re-initialises it
            'Following a Stop (state-loss), this SubclassProc won't be un-subclassed, so may be called
            'If there are orphaned timer messages hanging about, this should be able to ignore them without checking the TickerAPI
            'Also dereferencing will lead to errors as Stop destroys class definitions: https://stackoverflow.com/q/57560124/6609896
                If lParam <> 0 Then
                    Debug.Print printf("Rogue Timer!! ID:{0} TIMERPROC:{1}", wParam, lParam)
                    ManagedTimerMessageWindowSubclassProc = WinAPI.DefSubclassProc(hWnd, uMsg, wParam, lParam)
                    
                ElseIf TickerAPI.timerExists(wParam) Then 'there may be left-over messages in the queue which should be ignored by filtering for active timers
                    ManagedTimerMessageWindowSubclassProc = runTimerCallback(wParam)
                    
                Else
                    Debug.Print printf("Killing orphaned timer ID:{0} TIMERPROC:{1}", wParam, lParam)
                    On Error Resume Next
                    TickerAPI.KillTimerByID wParam
                    On Error GoTo 0
                    ManagedTimerMessageWindowSubclassProc = True 'skip it :)
                End If
            Else
                ManagedTimerMessageWindowSubclassProc = True 'skip
            End If
            
            
        Case Else
            ManagedTimerMessageWindowSubclassProc = WinAPI.DefSubclassProc(hWnd, uMsg, wParam, lParam)
            
    End Select
End Function

'TODO this should't be here
Private Function runTimerCallback(ByVal timerID As LongPtr) As LongPtr

    'timerID is a pointer to a wrapper; but since this is Friend code we can get it straight from the dictionary and avoid dereferencing
    On Error GoTo cleanFail 'catch unexpected error as we can't raise them here
    Dim callbackWrapper As ManagedCallbackWrapper
    If Not TickerAPI.tryGetManagedWrapper(timerID, callbackWrapper) Then
        Err.Raise 5, Description:=printf("Callback info for timerID:{0} not found", timerID) 'Raise error so we goto cleanFail
    End If
    
    'Bear in mind we only have the wrapper, not the actual timerProc; so get it from the wrapper interface
    Dim wrapperInterface As ICallbackWrapper
    Set wrapperInterface = callbackWrapper
        
    Dim timerProc As ITimerProc
    Set timerProc = wrapperInterface.Callback
    
    On Error Resume Next 'catch error from actually running the code
   
    runTimerCallback = timerProc.Exec(timerID, callbackWrapper.storedData)
    If Err.Number <> 0 Then logError "timerProc.Exec", Err.Number, Err.Description 'TODO print something specific to the timerProc
    On Error GoTo 0 'after check as it overwrites err.Number to 0
    
    Exit Function
    
cleanFail:
    logError "runTimerCallback", Err.Number, Err.Description
    runTimerCallback = False
    
End Function
