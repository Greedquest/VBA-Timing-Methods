Attribute VB_Name = "RecursionErrorExperiment"
'@Folder("Old.Tests")
'@IgnoreModule
Option Explicit
Private Const defaultTimerDelay As Long = 1000
Private Const MillisToSeconds As Double = 1 / 1000
Private Const defaultID As Long = 100
Private id As Long

Public hasBeenRun As Boolean

Private Declare Function ApiSetTimer Lib "user32" Alias "SetTimer" ( _
                         ByVal hWnd As Long, _
                         Optional ByVal nIDEvent As Long = defaultID, _
                         Optional ByVal uElapse As Long = defaultTimerDelay, _
                         Optional ByVal lpTimerFunc As Long) As Long

Private Declare Function ApiKillTimer Lib "user32" Alias "KillTimer" ( _
                         ByVal hWnd As Long, _
                         Optional ByVal nIDEvent As Long = defaultID) As Long
                         
'Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (ByVal lpMsg As LongPtr, ByVal hwnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Boolean

Private Sub testCallbackProc(ByVal windowHandle As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)

    Static timerSet As New Dictionary
    If Not timerSet.Exists(timerID) Then timerSet.Add timerID, 0
    timerSet(timerID) = timerSet(timerID) + 1
    
    Debug.Print windowHandle
    Debug.Print printf("Ticking - {0} (id:{1})", timerSet(timerID), timerID), time$
    If timerSet(timerID) > 10 Then
        On Error Resume Next
        TickerAPI.KillTimerByID timerID          'stop timer
    End If
    
End Sub

Public Sub quietSelfDestructingTimerProc(ByVal hWnd As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Debug.Print timerID; " Arrived"
    If timerID = defaultID Then
        ApiKillTimer hWnd, timerID
    Else
        ApiKillTimer 0, timerID
    End If
End Sub

Private Sub toggleDefaultTimer()
    Static runningID As Long
       
    If runningID = 0 Then
        ApiKillTimer Application.hWnd, defaultID
        runningID = -defaultID
    End If
    
    If runningID = -defaultID Then
        Debug.Print "Starting"
        hasBeenRun = False
        runningID = ApiSetTimer(Application.hWnd, lpTimerFunc:=AddressOf testCallbackProc)
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
    Dim msg As tagMSG
    toggleDefaultTimer
    tryPeekMessgageDelay msg, 2000
    DoEvents
    toggleDefaultTimer
End Sub

Public Sub testMessagePromotion()                'YEPPPPPP
    Dim quickTimer As Long, slowTimer As Long
    quickTimer = ApiSetTimer(0, 0, uElapse:=10, lpTimerFunc:=AddressOf quietSelfDestructingTimerProc)
    slowTimer = ApiSetTimer(Application.hWnd, uElapse:=1000, lpTimerFunc:=AddressOf quietSelfDestructingTimerProc)
    tightLoopDelay 1200                          'allow messages to build up
    
    Dim msg As tagMSG
    If Not tryPeekMessgageDelay(msg, 40, PM_NOREMOVE) Then
        Debug.Print "No message found"
        ApiKillTimer Application.hWnd, slowTimer
        ApiKillTimer 0, quickTimer
    End If
    Debug.Print printf("lParam: {0}, wParam: {1}", msg.lParam, msg.wParam)
End Sub

Public Sub doEventsDelay(Optional ByVal delayMillis As Long = defaultTimerDelay)
    Dim endTime As Single
    endTime = timer + delayMillis * MillisToSeconds
    Do While timer < endTime
        DoEvents
    Loop
End Sub

Public Function tryPeekMessgageDelay(outMsg As tagMSG, Optional ByVal delayMillis As Long = defaultTimerDelay, Optional ByVal flags As Long = PM_REMOVE, Optional ByVal timerID As Long = defaultID) As Boolean
    Dim endTime As Single
    endTime = timer + delayMillis * MillisToSeconds
    Do While timer < endTime
        If PeekMessage(outMsg, Application.hWnd, WM_TIMER, WM_TIMER, flags) <> 0 Then
            'Debug.Print printf("lParam: {0}, wParam: {1}", outMsg.lParam, outMsg.wParam)
            If outMsg.wParam = timerID Then      ' Or outMsg.lParam = 0 Then
                tryPeekMessgageDelay = True
                Exit Do
            End If
        End If
    Loop
End Function

Public Sub checkingForTimerMessageDelay(Optional ByVal delayMillis As Long = defaultTimerDelay)
    Dim endTime As Single
    endTime = timer + delayMillis * MillisToSeconds
    Do While timer < endTime
        Dim newStatus As DWORD
        newStatus = GetQueueStatus(QS_TIMER)
        If newStatus.HiWord <> 0 Then
            Dim reallyNewStatus As DWORD
            reallyNewStatus = GetQueueStatus(QS_TIMER)
            Debug.Print Toolbox.Strings.Format("New status: {0:x4} {1:x4} {2:x4} {3:x4}", newStatus.HiWord, newStatus.LoWord, reallyNewStatus.HiWord, reallyNewStatus.LoWord) '0010 0010 - 0010 0000
            Exit Do
        End If
        'DoEvents
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

Public Sub applicationOnTimeDelay(ByVal Callback As String, Optional ByVal delayMillis As Long = defaultTimerDelay)
    Application.OnTime futureTime(delayMillis), Callback
End Sub

Private Function futureTime(ByVal delayMillis As Long) As Variant
    futureTime = TimeSerial(Hour(Now), Minute(Now), Second(Now) + delayMillis * MillisToSeconds)
End Function


