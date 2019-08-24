VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SafeTickingTimerProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Callbacks")
Option Explicit

Implements ITimerProc

Private Function ITimerProc_Exec(ByVal timerID As Long, ByVal userData As Variant) As Boolean
    Static i As Long
    i = i + 1
    
    ITimerProc_Exec = True
    
    Dim expectedData As String
    On Error Resume Next
    expectedData = CStr(userData) 'catch error in case of bad data
    On Error GoTo 0
    
    Debug.Print Toolbox.Strings.Format("Ticking - {0} ({id:{1})\tData:'{3}'\t{2}", i, timerID, time$, expectedData)

    If i >= 10 Then
        TickerAPI.KillTimerByID timerID
    End If
    
End Function