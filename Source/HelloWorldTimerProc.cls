VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelloWorldTimerProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Tests.Experiments.Callbacks")

Implements ITimerProc

Private Sub ITimerProc_Exec(ByVal timerID As Long, ByVal userData As Variant, ByVal tickCount As Long)
    Err.Raise 5                                  'Debug.Print "Hey there matey"
End Sub

