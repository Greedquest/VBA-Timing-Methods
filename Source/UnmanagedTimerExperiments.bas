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

Public Sub testStopButton()
    TickerAPI.StartUnmanagedTimer AddressOf SafeCallbackProc, True, data:="User data!!"
    Debug.Print "Doing events"
    Dim endTime As Single: endTime = timer + 1
    Do While timer < endTime
        DoEvents
    Loop
    Debug.Print "Stopping"                       ', VBA.CLngPtr(AddressOf MessageWindowProcs.ManagedTimerMessageWindowSubclassProc)
    End
End Sub
