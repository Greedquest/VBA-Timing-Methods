Attribute VB_Name = "CallbackFunctions"

Option Explicit
'@Folder("Tests.Callbacks")

Public Sub SafeCallbackProc(ByRef createTimer As Bool, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Debug.Print "Callback called " & Time
    If message = WM_NOTIFY Then
        createTimer = False
    Else
        TickerAPI.KillTimerByID timerID
    End If
End Sub

Public Sub QuietTerminatingProc(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    If message = WM_NOTIFY Then
        Bool.FromPtr(createTimer) = False
    Else
        TickerAPI.KillTimerByID timerID
    End If
End Sub

Public Sub QuietNoOpCallback(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
End Sub

Public Sub SafeTickingProc(ByVal windowHandle As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Static i As Long
    Debug.Print "Ticking "; i
    i = i + 1
    If i > 10 Then TickerAPI.KillTimerByID timerID 'stop timer
End Sub

Public Sub RecursiveProc(ByRef createTimer As Bool, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Static i As Long
    i = i + 1
    Debug.Print i; "Callback called " & Time; timerID
    If i < 3 Then TickerAPI.StartTimer AddressOf RecursiveProc, True, 1000
    Debug.Print i
    i = i - 1
    createTimer = i = 1
End Sub

