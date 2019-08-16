Attribute VB_Name = "UnmanagedTimerExperiments"
'@Folder("Tests.Experiments")
Option Explicit

'Public Sub testSelfReplicating()
'    TickerAPI.StartUnmanagedTimer AddressOf RecursiveProc, True
'End Sub

Public Sub testImmediateTerminating()
    TickerAPI.UnlockApi
    TickerAPI.StartUnmanagedTimer AddressOf SafeCallbackProc, True
End Sub

Public Sub testAsyncTerminating()
    TickerAPI.UnlockApi
    TickerAPI.StartUnmanagedTimer AddressOf SafeCallbackProc, False
End Sub

Public Sub testImmediateTicking()
    TickerAPI.UnlockApi
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, True, data:="User data!!"
End Sub

Public Sub testAsyncTicking()
    TickerAPI.UnlockApi
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, False, data:="User data!!"
End Sub

Public Sub testInterwovenTicking()
    TickerAPI.UnlockApi
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, True, 1000, data:="Barry"
    doEventsDelay 500
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, True, 1000, data:="Suzie"
End Sub


