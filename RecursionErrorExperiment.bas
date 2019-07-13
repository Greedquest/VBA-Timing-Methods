Attribute VB_Name = "RecursionErrorExperiment"
'@Folder("Tests")
Option Explicit

Private Sub problematicSelfStartingCallback(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Static i As Long
    i = i + 1
    Debug.Print i; ": Self starter callback called - killing old timer"
    
    On Error GoTo checkError
    TickerAPI.KillTimerByID timerID
    On Error GoTo 0
    
continueAsNormal:
    Debug.Print i; ": Creating new timer"
    TickerAPI.StartTimer AddressOf problematicSelfStartingCallback, False, 1000
    Exit Sub
    
checkError:
    If Err.Number = TimerError.TimerNotFoundError Or Err.Number = TimerError.DestroyTimerError Then
        'do nothing
    Else
        Debug.Print Err.Number, Err.Description
    End If
    Resume continueAsNormal
        
End Sub


Sub testSelfStarter()
    TickerAPI.StartTimer AddressOf problematicSelfStartingCallback, False, 1000
End Sub
