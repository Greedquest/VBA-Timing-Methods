VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TickSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Decouples the TimerProc from metronome or any other second level API objects to avoid circular references (since TickerAPI holds a reference to the ITimerProc)"
'@Folder("SecondLevelAPI.Common")
'@ModuleDescription("Decouples the TimerProc from metronome or any other second level API objects to avoid circular references (since TickerAPI holds a reference to the ITimerProc)")
Option Explicit

Public Event Tick(ByVal timerID As LongPtr, ByVal userData As Variant, ByVal tickCount As Long)

Implements ITimerProc

Private Sub ITimerProc_Exec(ByVal timerID As LongPtr, ByVal userData As Variant, ByVal tickCount As Long)
    RaiseEvent Tick(timerID, userData, tickCount)
End Sub
