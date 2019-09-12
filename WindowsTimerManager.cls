VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindowsTimerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FirstLevelAPI")
Option Explicit

Implements ITimerManager

Private Sub ITimerManager_KillTimer(ByVal data As TimerData)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Sub ITimerManager_StartTimer(ByVal data As TimerData)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Sub ITimerManager_UpdateTimer(ByVal data As TimerData)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Sub ITimerManager_TriggerTimer(ByVal data As TimerData)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Sub ITimerManager_KillAllTimers(ByRef dataArray() As TimerData)
    Err.Raise 5 'TODO implement interface member
End Sub

