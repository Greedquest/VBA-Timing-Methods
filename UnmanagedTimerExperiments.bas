Attribute VB_Name = "UnmanagedTimerExperiments"
'@Folder("Tests.Experiments")
Option Explicit

'Public Sub testSelfReplicating()
'    TickerAPI.StartUnmanagedTimer AddressOf RecursiveProc, True
'End Sub

Public Sub testImmediateTerminating()
    TickerAPI.StartUnmanagedTimer AddressOf SafeCallbackProc, True
End Sub

Public Sub testAsyncTerminating()                'TickerAPI.killalltimers
    TickerAPI.StartUnmanagedTimer AddressOf SafeCallbackProc, False
End Sub

Public Sub testImmediateTicking()
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, True
End Sub

Public Sub testAsyncTicking()
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, False
End Sub

Public Sub testInterwovenTicking()
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, True, 1000
    doEventsDelay 500
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, True, 1000
End Sub


