Attribute VB_Name = "WaitableCallback"
'@Folder("SecondLevelAPI")
Option Explicit

Public Const InfiniteTicks As Long = -1

Public Sub WaitableTimerCallbackProc(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    'Process message and forward to parent
    MetronomeCollection.CallbackNotify createTimer, message, timerID, tickCount
End Sub
