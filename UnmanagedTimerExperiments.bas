Attribute VB_Name = "UnmanagedTimerExperiments"
'@Folder("Tests.Experiments")
Option Explicit

'Public Sub testSelfReplicating()
'    TickerAPI.StartUnmanagedTimer AddressOf RecursiveProc, True
'End Sub

Public Sub testImmediateTerminating()
    TickerAPI.StartUnmanagedTimer AddressOf SafeCallbackProc, True, data:="User data!!"
End Sub

Public Sub testAsyncTerminating()
    TickerAPI.StartUnmanagedTimer AddressOf SafeCallbackProc, False, data:="User data!!"
End Sub

Public Sub testImmediateTicking()
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, True, data:="User data!!"
End Sub

Public Sub testAsyncTicking()
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, False, data:="User data!!"
End Sub

Public Sub testInterwovenTicking()
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, True, 1000, data:="Barry"
    doEventsDelay 500
    TickerAPI.StartUnmanagedTimer AddressOf terminatingIndexedTickingProc, True, 1000, data:="Suzie"
End Sub
