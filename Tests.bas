Attribute VB_Name = "Tests"
'@Folder("Tests.Stubs")
Option Explicit

'
'Private Declare Function SetTimer Lib "user32" ( _
'                         ByVal HWnd As Long, _
'                         ByVal nIDEvent As Long, _
'                         ByVal uElapse As Long, _
'                         ByVal lpTimerFunc As Long) As Long
'
'Public Declare Function killTimer Lib "user32" Alias _
'                         "KillTimer" (ByVal HWnd As _
'                         Long, ByVal nIDEvent As Long) As Long
'
'Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
'                         ByVal lpPrevWndFunc As Long, _
'                         ByRef HWnd As Boolean, _
'                         ByVal msg As Long, _
'                         ByVal wParam As Long, _
'                         ByVal lParam As Long) As Long
Public Sub testInfiniteRecursion()
    'TODO Bug infinite recursion can't be stopped as timer messages hang about even when timers are killed - flush message queue or otherwise prevent timers being made during recursion
    TickerAPI.StartTimer AddressOf RecursiveProc, True
End Sub

Public Sub testSyncTerminating()
    TickerAPI.StartTimer AddressOf SafeCallbackProc, True
End Sub

Public Sub testAsyncTerminating()                'TickerAPI.killalltimers
    TickerAPI.StartTimer AddressOf SafeCallbackProc, False
End Sub

Public Sub testSyncTicking()
    TickerAPI.StartTimer AddressOf terminatingIndexedTickingProc, True
End Sub

Public Sub testAsyncTicking()
    TickerAPI.StartTimer AddressOf terminatingIndexedTickingProc, False
End Sub

Public Sub testInterwovenTicking()
    TickerAPI.StartTimer AddressOf terminatingIndexedTickingProc, True, 1000
    doEventsDelay 500
    TickerAPI.StartTimer AddressOf terminatingIndexedTickingProc, True, 1000
End Sub


