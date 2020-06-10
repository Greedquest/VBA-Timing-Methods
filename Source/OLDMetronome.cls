VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OLDMetronome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("OLDSecondLevelAPI")
'@IgnoreModule
Option Explicit
Implements OLDIMetronomeControls

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

Friend Function Controls() As OLDIMetronomeControls 'TODO decide scope
    Set Controls = Me
End Function

Private Property Get OLDIMetronomeControls_isRunning() As Boolean
    OLDIMetronomeControls_isRunning = this.isRunning
End Property

Private Property Let OLDIMetronomeControls_isRunning(ByVal value As Boolean)
    If this.isRunning <> value Then              'only do stuff if it would change things
        If value = True Then
            If OLDMetronomeCollection.tryInitialiseTimer(Me) Then this.isRunning = True
        Else
            If OLDMetronomeCollection.tryTerminateTimer(Me) Then this.isRunning = False
        End If
    End If
End Property

Private Property Get OLDIMetronomeControls_isFirstTickSynchronous() As Boolean
    OLDIMetronomeControls_isFirstTickSynchronous = this.isFirstTickSynchronous
End Property

Private Property Let OLDIMetronomeControls_isFirstTickSynchronous(ByVal value As Boolean)
    this.isFirstTickSynchronous = value
End Property

Private Property Get OLDIMetronomeControls_timerID() As Long
    OLDIMetronomeControls_timerID = this.timerID
End Property

Private Property Let OLDIMetronomeControls_timerID(ByVal value As Long)
    this.timerID = value
End Property

Private Property Get OLDIMetronomeControls_remainingTicks() As Long
    OLDIMetronomeControls_remainingTicks = this.remainingTicks
End Property

Private Property Let OLDIMetronomeControls_remainingTicks(ByVal value As Long)
    this.remainingTicks = value
End Property

Private Property Get OLDIMetronomeControls_tickIntervalMillis() As Long
    OLDIMetronomeControls_tickIntervalMillis = this.tickIntervalMillis
End Property

Private Property Let OLDIMetronomeControls_tickIntervalMillis(ByVal value As Long)
    this.tickIntervalMillis = value
End Property

Private Sub OLDIMetronomeControls_Tick()
    RaiseEvent Tick
End Sub

