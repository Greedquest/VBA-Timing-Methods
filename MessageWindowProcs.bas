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
    'TODO assumes timerID is the raw callback, instead get an IFunction and run it in On Error guard
    'TODO think about params
    runTimerCallback = WinAPI.CallWindowProc(timerID, Bool.Create(False), WM_TIMER)
End Function
