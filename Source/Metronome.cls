VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Metronome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SecondLevelAPI")
'@PredeclaredID
Option Explicit

Public Event Tick()

Private Type TMetronome
    ID As LongPtr
    TickOnStart As Boolean
    MinimumTickIntervalMillis As Long
    isPaused As Boolean
End Type

Private this As TMetronome
Private WithEvents ticker As EventRaisingTimerProc
Attribute ticker.VB_VarHelpID = -1

Public Function Create(ByVal tickFrequencyMillis As Long, Optional ByVal delayedStart As Boolean = False, Optional ByVal firstTickOnStart As Boolean = True) As Metronome
    With New Metronome
        Set Create = .Self
        .TickOnStart = firstTickOnStart
        .MinimumTickIntervalMillis = tickFrequencyMillis
        If Not delayedStart Then .Start
    End With
End Function

Friend Property Get Self() As Metronome
    Set Self = Me
End Property

Public Property Get TickOnStart() As Boolean
    TickOnStart = this.TickOnStart
End Property

Public Property Let TickOnStart(ByVal value As Boolean)
    this.TickOnStart = value
End Property

Public Property Get MinimumTickIntervalMillis() As Long
    MinimumTickIntervalMillis = this.MinimumTickIntervalMillis
End Property

Public Property Let MinimumTickIntervalMillis(ByVal RHS As Long)
    this.MinimumTickIntervalMillis = RHS
End Property

Public Sub Start()
    Static hasBeenStarted As Boolean
    If Not hasBeenStarted Then                   'Should only call Start once, since there should only ever be one api timer per metronome
        this.ID = TickerAPI.StartManagedTimer(TimerProc, TickOnStart, MinimumTickIntervalMillis)
        hasBeenStarted = True
    End If
End Sub

Private Property Get TimerProc() As ITimerProc
    If ticker Is Nothing Then Set ticker = New EventRaisingTimerProc
    Set TimerProc = ticker
End Property

'@Description("Temporarily pauses ticker by ignoring the messages. To kill the timer completely, just destroy the metronome object")
Public Property Let Paused(ByVal isPaused As Boolean)
Attribute Paused.VB_Description = "Temporarily pauses ticker by ignoring the messages. To kill the timer completely, just destroy the metronome object"
    this.isPaused = True
End Property

Public Property Get Paused() As Boolean
    Paused = this.isPaused
End Property

Private Sub Class_Terminate() 'BUG never called since ticker API holds a strong reference to this object (for the calling stuff)
    TickerAPI.KillTimerByID this.ID
End Sub

Private Sub ticker_Tick(ByVal timerID As Long, ByVal userData As Variant, ByVal tickCount As Long)
    If Not Paused Then RaiseEvent Tick
End Sub
