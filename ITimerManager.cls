VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITimerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FirstLevelAPI")
'@Interface
Option Explicit

Public Sub KillTimer(ByVal data As TimerData)
End Sub

Public Sub StartTimer(ByVal data As TimerData)
End Sub

Public Sub UpdateTimer(ByVal data As TimerData)
End Sub

Public Sub TriggerTimer(ByVal data As TimerData)
End Sub

Public Sub KillAllTimers(ByRef dataArray() As TimerData)
End Sub
