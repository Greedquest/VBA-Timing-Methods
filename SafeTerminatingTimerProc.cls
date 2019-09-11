VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SafeTerminatingTimerProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Experiments.Callbacks")
Option Explicit

Implements ITimerProc

'@Ignore ParameterNotUsed: callbacks need to have this signature regardless
Private Sub ITimerProc_Exec(ByVal timerID As Long, ByVal userData As Variant, ByVal tickCount As Long)
    Debug.Print "Callback called", time$, userData
    TickerAPI.KillTimerByID timerID
End Sub
