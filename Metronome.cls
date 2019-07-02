VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Metronome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TimerAPI")
Option Explicit
Implements IMetronomeControls

Private Type TMetronome
    isRunning As Boolean
    isFirstTickSynchronous As Boolean
    timerID As Long
    remainingTicks As Long
    tickIntervalMillis As Long
End Type

Private this As TMetronome

Public Event Tick()

Public Sub Start()
    
End Sub

Public Sub Pause()

End Sub

Public Function Controls() As IMetronomeControls
    Set Controls = Me
End Function


Private Property Get IMetronomeControls_isRunning() As Boolean
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Let IMetronomeControls_isRunning(ByVal value As Boolean)
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Get IMetronomeControls_isFirstTickSynchronous() As Boolean
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Let IMetronomeControls_isFirstTickSynchronous(ByVal value As Boolean)
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Get IMetronomeControls_timerID() As Long
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Let IMetronomeControls_timerID(ByVal value As Long)
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Get IMetronomeControls_remainingTicks() As Long
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Let IMetronomeControls_remainingTicks(ByVal value As Long)
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Get IMetronomeControls_tickIntervalMillis() As Long
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Let IMetronomeControls_tickIntervalMillis(ByVal value As Long)
    Err.Raise 5 'TODO implement interface member
End Property

Private Sub IMetronomeControls_Tick()
    Err.Raise 5 'TODO implement interface member
End Sub

