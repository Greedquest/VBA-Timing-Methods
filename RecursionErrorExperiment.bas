Attribute VB_Name = "RecursionErrorExperiment"
'@Folder("Tests")
Option Explicit
Private Const defaultTimerDelay As Long = 1000
Private Const MillisToSeconds As Double = 1 / 1000
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

Sub testVariousMessageDelays()
'Possible delay sources:
' - DoEvents Loop
' - Application.Wait (synchronous sleep)
' - Api Sleep / Application.OnTime (Async sleep)
' - Actual work; tight timer loop
' - UI work; calculation / edit cells

'QUs:
'Do messages ever build up
' - from one/ multiple sources
' - with different delays

End Sub

Public Sub doEventsDelay(Optional ByVal delayMillis As Long = defaultTimerDelay)
    Dim endTime As Single
    endTime = timer + delayMillis
    Do While timer < endTime
        DoEvents
    Loop
End Sub

Public Sub applicationWaitDelay(Optional ByVal delayMillis As Long = defaultTimerDelay)
    Application.Wait futureTime(delayMillis)
End Sub

Public Sub tightLoopDelay(Optional ByVal delayMillis As Long = defaultTimerDelay)
    Dim endTime As Single
    endTime = timer + delayMillis
    Do While timer < endTime
        'do nothing :(
    Loop
End Sub

Public Sub applicationOnTimeDelay(ByVal callback As String, Optional ByVal delayMillis As Long = defaultTimerDelay)
    Application.OnTime futureTime(delayMillis), callback
End Sub

Private Function futureTime(ByVal delayMillis As Long) As Variant
    futureTime = TimeSerial(Hour(Now), Minute(Now), Second(Now) + delayMillis * MillisToSeconds)
End Function
