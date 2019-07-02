Attribute VB_Name = "WaitableCallback"
'@Folder("TimerAPI")
Option Explicit

Public Sub WaitableTimerCallbackProc(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
'Process message and forward to parent
    WaitableTimers.CallbackNotify createTimer, message, timerID, tickCount
End Sub
