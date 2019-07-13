Attribute VB_Name = "RecursionErrorExperiment"
'@Folder("Tests")
Option Explicit

Private Sub problematicSelfStartingCallback(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Static i As Long
    i = i + 1
    Debug.Print i; ": Self starter callback called"
    TickerAPI.StartTimer AddressOf problematicSelfStartingCallback, False, 1000
End Sub


Sub testSelfStarter()
    TickerAPI.StartTimer AddressOf problematicSelfStartingCallback, False, 1000
End Sub
