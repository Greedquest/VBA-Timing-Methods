Attribute VB_Name = "RecursionErrorExperiment"
'@Folder("Tests")
Option Explicit
Private Const timerDelay As Long = 2000
Private id As Long

Private Declare Function ApiSetTimer Lib "user32" Alias "SetTimer" ( _
                         ByVal HWnd As Long, _
                         ByVal nIDEvent As Long, _
                         ByVal uElapse As Long, _
                         ByVal lpTimerFunc As Long) As Long

Private Declare Function ApiKillTimer Lib "user32" Alias "KillTimer" ( _
                         ByVal HWnd As Long, _
                         ByVal nIDEvent As Long) As Long

Sub toggleTimer()
    Static runningID As Long
    Const defaultID As Long = 100
    
    If runningID = 0 Then
        ApiKillTimer Application.HWnd, defaultID
        runningID = -defaultID
    End If
    
    If runningID = -defaultID Then
        Debug.Print "Starting"
        runningID = ApiSetTimer(Application.HWnd, defaultID, timerDelay, AddressOf CallbackFunctions.SafeTickingProc)
    Else
        Debug.Print "Stopping"
        ApiKillTimer Application.HWnd, defaultID
        runningID = -defaultID
    End If
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
