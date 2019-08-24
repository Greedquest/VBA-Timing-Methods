VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SafeTerminatingTimerProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Callbacks")
Option Explicit

Implements ITimerProc

Private Function ITimerProc_Exec(ByVal timerID As LongPtr, ByVal userData As Variant) As Boolean
    Debug.Print "Callback called", time$, userData
    TickerAPI.KillTimerByID timerID
End Function

