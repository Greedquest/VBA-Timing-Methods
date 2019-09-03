VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Metronome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("SecondLevelAPI")
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
    Controls.isRunning = True
End Sub

Public Sub Pause()
    Controls.isRunning = False
End Sub

Friend Function Controls() As IMetronomeControls 'TODO decide scope
    Set Controls = Me
End Function

Private Property Get IMetronomeControls_isRunning() As Boolean
    IMetronomeControls_isRunning = this.isRunning
End Property

Private Property Let IMetronomeControls_isRunning(ByVal value As Boolean)
    If this.isRunning <> value Then              'only do stuff if it would change things
        If value = True Then
            If MetronomeCollection.tryInitialiseTimer(Me) Then this.isRunning = True
        Else
            If MetronomeCollection.tryTerminateTimer(Me) Then this.isRunning = False
        End If
    End If
End Property

Private Property Get IMetronomeControls_isFirstTickSynchronous() As Boolean
    IMetronomeControls_isFirstTickSynchronous = this.isFirstTickSynchronous
End Property

Private Property Let IMetronomeControls_isFirstTickSynchronous(ByVal value As Boolean)
    this.isFirstTickSynchronous = value
End Property

Private Property Get IMetronomeControls_timerID() As Long
    IMetronomeControls_timerID = this.timerID
End Property

Private Property Let IMetronomeControls_timerID(ByVal value As Long)
    this.timerID = value
End Property

Private Property Get IMetronomeControls_remainingTicks() As Long
    IMetronomeControls_remainingTicks = this.remainingTicks
End Property

Private Property Let IMetronomeControls_remainingTicks(ByVal value As Long)
    this.remainingTicks = value
End Property

Private Property Get IMetronomeControls_tickIntervalMillis() As Long
    IMetronomeControls_tickIntervalMillis = this.tickIntervalMillis
End Property

Private Property Let IMetronomeControls_tickIntervalMillis(ByVal value As Long)
    this.tickIntervalMillis = value
End Property

Private Sub IMetronomeControls_Tick()
    RaiseEvent Tick
End Sub
