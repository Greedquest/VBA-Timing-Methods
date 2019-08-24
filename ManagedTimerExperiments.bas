Attribute VB_Name = "ManagedTimerExperiments"
'@Folder("Tests.Experiments")
Option Explicit

Public Sub testImmediateTerminating()
    Err.Raise 5
End Sub

Public Sub testAsyncTerminating()
    TickerAPI.StartUnmanagedTimer AddressOf SafeCallbackProc, False, data:="User data!!"
End Sub

Public Sub testImmediateTicking()
    TickerAPI.StartManagedTimer New SafeTickingTimerProc, True, data:="User data!!"
End Sub

Public Sub testAsyncTicking()
    TickerAPI.StartManagedTimer New SafeTickingTimerProc, False, data:="User data!!"
End Sub

Public Sub testInterwovenTicking()
    TickerAPI.StartManagedTimer New SafeTickingTimerProc, True, 1000, data:="Barry"
    doEventsDelay 500
    TickerAPI.StartManagedTimer New SafeTickingTimerProc, True, 1000, data:="Suzie"
End Sub
