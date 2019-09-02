VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITimerProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("FirstLevelAPI.Utils.Wrappers")
'@PredeclaredID
'@Exposed
'@Interface
Option Explicit

'Interface
Public Sub Exec(ByVal timerID As LongPtr, ByVal userData As Variant, ByVal tickCount As Long)
    Err.Raise 5                                  'not implemented
End Sub

