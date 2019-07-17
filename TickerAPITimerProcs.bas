Attribute VB_Name = "TickerAPITimerProcs"
'@Folder("API.Utils")
Option Explicit

Public Sub UnlockCallbackProc(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
'TODO: this is public so try and catch fake calls
    TickerAPI.UnlockApi
    If message = WM_NOTIFY Then
        Bool.FromPtr(createTimer) = False
    Else
        TickerAPI.KillTimerByID timerID
    End If
End Sub
