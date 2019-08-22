Attribute VB_Name = "MessageWindowProcs"
'@Folder("FirstLevelAPI.Utils")
Option Explicit
Option Private Module

Public Function ManagedTimerMessageWindowSubclassProc(ByVal hWnd As LongPtr, ByVal uMsg As WindowsMessage, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    
    Debug.Print "MSG #", uMsg
    Select Case uMsg
    
            'NOTE this will never receive timer messages where TIMERPROC is specified,
        Case WindowsMessage.WM_TIMER             'wParam = timerID , lParam = "timerProc" (will be Null if it reaches here)
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
            
        Case Else
            ManagedTimerMessageWindowSubclassProc = WinAPI.DefSubclassProc(hWnd, uMsg, wParam, lParam)
            
    End Select
End Function

'TODO this should't be here
Private Function runTimerCallback(ByVal timerID As LongPtr) As LongPtr
    On Error Resume Next
    runTimerCallback = ITimerProc.FromPtr(timerID).Exec
    If Err.Number <> 0 Then logError "runTimerCallback", Err.Number, Err.Description
    On Error GoTo 0
End Function

Private Sub logError(ByVal Source As String, ByVal errNum As Long, ByVal errDescription As String)
    If Not LogManager.IsEnabled(ErrorLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("TickerAPI", ErrorLevel)
    End If
    LogManager.log ErrorLevel, Toolbox.Strings.Format("{0} raised an error: #{1} - {2}", Source, errNum, errDescription)
End Sub

