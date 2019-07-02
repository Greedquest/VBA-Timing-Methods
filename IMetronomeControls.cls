VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMetronomeControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TimerAPI")
Option Explicit

Public isRunning As Boolean
Public isFirstTickSynchronous As Boolean
Public timerID As Long
Public remainingTicks As Long
Public tickIntervalMillis As Long

Public Sub Tick()

End Sub
