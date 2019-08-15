Attribute VB_Name = "TickerAPITimerProcs"
'@Folder("FirstLevelAPI.Utils")
Option Explicit
Option Private Module

Public Sub UnlockCallbackProc(ByVal hWnd As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    'TODO: this is public so try and catch fake calls
    TickerAPI.UnlockApi
    On Error Resume Next                         'if this is floating around then it should be hoovered up by the api anyway so no point raising expected errors
    TickerAPI.KillTimerByID timerID
    On Error GoTo 0
End Sub

