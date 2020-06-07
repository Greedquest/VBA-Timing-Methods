VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleMetronome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SecondLevelAPI")
'@PredeclaredID

Option Explicit

Public Event Tick(ByVal timerID As Long, ByVal userData As Variant, ByVal tickCount As Long)

Private Type MetronomeData
    tickMillis As Long
    isTicking As Boolean
    timerID As LongPtr
End Type

Implements ITimerProc

Private this As MetronomeData

Private Sub Class_Initialize()
    tickFrequency = 500 'set defaults
End Sub
    
Public Property Let tickFrequency(ByVal tickMillis As Long)
    this.tickMillis = tickMillis
    If this.isTicking Then UpdateTimer
End Property

Public Property Get tickFrequency() As Long
    tickFrequency = this.tickMillis
End Property

Public Sub StartTicking(ByVal tickImmediately As Boolean, ByVal userData As Variant)
    this.timerID = TickerAPI.StartManagedTimer(Me, tickImmediately, tickFrequency, userData)
End Sub

Public Sub StopTicking()
    TickerAPI.KillTimerByID (this.timerID)
End Sub

Private Sub UpdateTimer()
    StopTicking
    StartTicking
End Sub

Private Sub ITimerProc_Exec(ByVal timerID As Long, ByVal userData As Variant, ByVal tickCount As Long)
    RaiseEvent Tick(timerID, userData, tickCount)
End Sub
