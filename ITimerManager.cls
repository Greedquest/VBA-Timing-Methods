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

Public Enum InternalTimerError
    [_ErrBase] = 6                          'just in case of clashes, let's offset the errors
    [_Start] = vbObjectError + [_ErrBase]   'TimerError.[_End] - 1
    CreateMessageWindowError
    APIKillTimerError
    CastKeyToWrapperError
    APIStartTimerError
    APIPostMessageError
End Enum

Public Sub KillTimer(ByVal data As TimerData)
End Sub

Public Sub StartTimer(ByVal data As TimerData, ByVal delayMillis As Long)
End Sub

Public Sub UpdateTimer(ByVal data As TimerData, ByVal delayMillis As Long)
End Sub

Public Function TryTriggerTimer(ByVal data As TimerData) As Boolean
End Function

Public Sub KillAllTimers(ByVal dataArray As Variant)
End Sub
