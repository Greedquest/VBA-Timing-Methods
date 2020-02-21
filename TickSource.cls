VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TickSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SecondLevelAPI")
Option Explicit

Public Event Tick()

Private Enum MetronomeState
    msRunning
    msPaused
    msStopped
End Enum

Public Sub PauseTicking()

End Sub

Public Sub ResumeTicking()

End Sub

Public Sub StopTicking()

End Sub

