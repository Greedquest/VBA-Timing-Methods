Attribute VB_Name = "RecursionErrorExperiment"
'@Folder("Tests")
Option Explicit
Private Const timerDelay As Long = 2000
Private id As Long

Public Sub immediateCallback(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)

End Sub

Private Sub problematicSelfStartingCallback(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Static i As Long
    i = i + 1
    Debug.Print i; ": Self starter callback called - killing old timer"
    
    On Error GoTo checkError
    TickerAPI.KillTimerByID timerID
    On Error GoTo 0
    
continueAsNormal:
    If i < 5 Then
        Debug.Print i; ": Creating new timer"
        TickerAPI.StartTimer AddressOf problematicSelfStartingCallback, False, timerDelay
    Else
        Debug.Print i; ": Recursion limit reached"
    End If
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
    Debug.Print "hi"
    id = TickerAPI.StartTimer(AddressOf problematicSelfStartingCallback, False, timerDelay)
    Debug.Print "ho"
    'Application.Wait TimeSerial(Hour(Now), Minute(Now), Second(Now) + 8)

End Sub

Sub endSelfStarter()
    Debug.Print "silver lining"
    On Error Resume Next
    TickerAPI.KillTimerByID id
End Sub
