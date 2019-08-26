Attribute VB_Name = "Delays"
'@Folder("Tests.Utils")
Option Explicit
Private Const defaultTimerDelay As Long = 1000
Private Const MillisToSeconds As Double = 1 / 1000
Private Const defaultID As Long = 100


Public Sub doEventsDelay(Optional ByVal delayMillis As Long = defaultTimerDelay)
    Dim endTime As Single
    endTime = timer + delayMillis * MillisToSeconds
    Do While timer < endTime
        DoEvents
    Loop
End Sub

Public Function tryPeekMessgageDelay(outMsg As WinAPI.timerMessage, Optional ByVal delayMillis As Long = defaultTimerDelay, Optional ByVal flags As Long = PM_REMOVE, Optional ByVal timerID As Long = defaultID) As Boolean
    Dim endTime As Single
    endTime = timer + delayMillis * MillisToSeconds
    Do While timer < endTime
        If PeekTimerMessage(outMsg, TickerAPI.messageWindowHandle, WM_TIMER, WM_TIMER, flags) <> 0 Then
            'Debug.Print printf("lParam: {0}, wParam: {1}", outMsg.lParam, outMsg.wParam)
            If outMsg.timerID = timerID Then      ' Or outMsg.lParam = 0 Then
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
