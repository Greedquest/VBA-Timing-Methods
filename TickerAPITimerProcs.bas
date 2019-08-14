Attribute VB_Name = "TickerAPITimerProcs"
'@Folder("FirstLevelAPI.Utils")
Option Explicit
Option Private Module

Public Sub UnlockCallbackProc(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    'TODO: this is public so try and catch fake calls
    TickerAPI.UnlockApi
    If message = WM_NOTIFY Then                  'should never be true as this is meant to be called async
        TickerParams.FromPtr(createTimer).TickerIsRunning = False
    Else
        On Error Resume Next                     'if this is floating around then it should be hoovered up by the api anyway so no point raising expected errors
        TickerAPI.KillTimerByID timerID
        On Error GoTo 0
    End If
End Sub

