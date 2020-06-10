Attribute VB_Name = "ManagedTimerExperiments"
'@Folder("Tests.Experiments")
Option Explicit

Public Sub testImmediateTerminating()
    TickerAPI.StartManagedTimer New SafeTerminatingTimerProc, True, data:="User data!!"
End Sub

Public Sub testAsyncTerminating()
    TickerAPI.StartManagedTimer New SafeTerminatingTimerProc, False, data:="User data!!"
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

