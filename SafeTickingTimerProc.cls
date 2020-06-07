VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SafeTickingTimerProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Experiments.Callbacks")
Option Explicit

Implements ITimerProc

'@Ignore ParameterNotUsed: callbacks need to have this signature regardless
Private Sub ITimerProc_Exec(ByVal timerID As Long, ByVal userData As Variant, ByVal tickCount As Long)
    Static I As Long
    I = I + 1
       
    Dim expectedData As String
    On Error Resume Next
    expectedData = CStr(userData)                'catch error in case of bad data
    On Error GoTo 0
    
    Debug.Print printf("Ticking - {0} (id:{1})    Data:'{3}'    {2}", I, timerID, time$, expectedData)

    If I >= 10 Then
        TickerAPI.KillTimerByID timerID
    End If
End Sub
