VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OLDIMetronomeControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("OLDSecondLevelAPI")
'@IgnoreModule
Option Explicit

'@IgnoreModule EncapsulatePublicField
Public isRunning As Boolean
Public isFirstTickSynchronous As Boolean
Public timerID As Long
Public remainingTicks As Long
Public tickIntervalMillis As Long

Public Sub Tick()
End Sub

