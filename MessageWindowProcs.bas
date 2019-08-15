Attribute VB_Name = "MessageWindowProcs"
'@Folder("FirstLevelAPI.Utils")
Option Explicit
Option Private Module

Public Function MessageWindowSubclassProc(ByVal hWnd As LongPtr, ByVal uMsg As WindowsMessage, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    
    Debug.Print "Message #"; uMsg
    Select Case uMsg
    
        'NOTE this will never receive timer messages where TIMERPROC is specified,
        Case WindowsMessage.WM_TIMER 'wParam = timerID , lParam = "timerProc" (will be Null if it reaches here)
            If TickerAPI.timerExists(wParam) Then
                MessageWindowSubclassProc = runTimerCallback(wParam)  'WinAPI.DefSubclassProc(hwnd, uMsg, wParam, lParam)
            Else
                On Error Resume Next 'checking for the timer should trigger destruction
                TickerAPI.KillTimerByID wParam
                On Error GoTo 0
                MessageWindowSubclassProc = True 'skip it :)
            End If
            
        Case Else
            MessageWindowSubclassProc = WinAPI.DefSubclassProc(hWnd, uMsg, wParam, lParam)
            
    End Select
End Function

Private Function runTimerCallback(ByVal timerID As LongPtr) As LongPtr
    On Error Resume Next
        runTimerCallback = ITimerProc.FromPtr(timerID).Exec
        If Err.Number <> 0 Then logError "runTimerCallback", Err.Number, Err.Description
    On Error GoTo 0
End Function

Private Sub logError(ByVal source As String, ByVal errNum As Long, ByVal errDescription As String)
    If Not LogManager.IsEnabled(ErrorLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("TickerAPI", ErrorLevel)
    End If
    LogManager.log ErrorLevel, Toolbox.Strings.Format("{0} raised an error: #{1} - {2}", source, errNum, errDescription)
End Sub
