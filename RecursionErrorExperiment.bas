Attribute VB_Name = "RecursionErrorExperiment"
'@Folder("Tests")
Option Explicit
Private Const defaultTimerDelay As Long = 1000
Private Const MillisToSeconds As Double = 1 / 1000
Private Const defaultID As Long = 100
Private id As Long

Private Declare Function ApiSetTimer Lib "user32" Alias "SetTimer" ( _
                         ByVal hWnd As Long, _
                         Optional ByVal nIDEvent As Long = defaultID, _
                         Optional ByVal uElapse As Long = defaultTimerDelay, _
                         Optional ByVal lpTimerFunc As Long) As Long

Private Declare Function ApiKillTimer Lib "user32" Alias "KillTimer" ( _
                         ByVal hWnd As Long, _
                         ByVal nIDEvent As Long) As Long
                         
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (ByVal lpMsg As LongPtr, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Boolean
                        
Private Sub toggleDefaultTimer()
    Static runningID As Long
       
    If runningID = 0 Then
        ApiKillTimer Application.hWnd, defaultID
        runningID = -defaultID
    End If
    
    If runningID = -defaultID Then
        Debug.Print "Starting"
        runningID = ApiSetTimer(Application.hWnd, lpTimerFunc:=AddressOf CallbackFunctions.SafeTickingProc)
    Else
        Debug.Print "Stopping"
        ApiKillTimer Application.hWnd, defaultID
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

Sub testOneSource()
    'TickerAPI.StartTimer AddressOf SafeTickingProc
End Sub

Public Sub doEventsDelay(Optional ByVal delayMillis As Long = defaultTimerDelay)
    Dim endTime As Single
    endTime = timer + delayMillis * MillisToSeconds
    Do While timer < endTime
        DoEvents
    Loop
End Sub

Public Sub applicationWaitDelay(Optional ByVal delayMillis As Long = defaultTimerDelay)
    Application.Wait futureTime(delayMillis)
End Sub

Public Sub tightLoopDelay(Optional ByVal delayMillis As Long = defaultTimerDelay)
    Dim endTime As Single
    endTime = timer + delayMillis * MillisToSeconds
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
