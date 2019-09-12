VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindowsTimerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FirstLevelAPI")
Option Explicit

Implements ITimerManager

Private Type windowsTimerManagerData
    messageWindow As New ModelessMessageWindow
End Type

Private this As windowsTimerManagerData

Private Sub ITimerManager_KillTimer(ByVal data As TimerData)
    'NOTE no need to clear messages as killing the timer invalidates any which have a TIMERPROC argument (which they all do)
    On Error GoTo cleanFail
    
    '0 indicates some failure
    If WinAPI.KillTimer(this.messageWindow.handle, data.id) = 0 Then
        throwDllError Err.LastDllError, "Call returned zero, probably tried to kill non-existent timer"
    End If
    
cleanExit:
    'TODO If timerExists(data.ID) Then this.timerIDs.Remove da 'remove regardless of success, since it's either gone now, or was invalid anyway
    Exit Sub
    
cleanFail:
    logError "WinAPI.KillTimer", Err.Number, Err.Description
    raiseInternalError APIKillTimerError, "KillTimer"
    Resume cleanExit
End Sub

Private Sub ITimerManager_StartTimer(ByVal data As TimerData, ByVal delayMillis As Long)
    Const loggerSourceName As String = "StartTimer"
   
    On Error GoTo setTimerFail                   'Custom handler so we can log precise dll errors and condense error messages + clear up any timer which may have been made
    
    Dim newTimerID As LongPtr
    newTimerID = WinAPI.SetTimer(this.messageWindow.handle, data.callbackWrapperPointer, delayMillis, data.timerProc)
    
    If newTimerID = 0 Then
        throwDllError Err.LastDllError
    
    ElseIf newTimerID <> data.callbackWrapperPointer Then
        Err.Raise 5, Description:="timerID does not have expected value" 'REVIEW is there a better assertion error to raise?
        
    End If
    
    data.id = newTimerID
    Exit Sub
    
setTimerFail:
    logError "WinAPI.SetTimer", Err.Number, Err.Description
    KillTimer windowHandle, newTimerID
    raiseInternalError CastKeyToWrapperError, loggerSourceName
    Resume                                       'for debugging - break above and jump to the error-raising statement

End Sub

Private Sub ITimerManager_UpdateTimer(ByVal data As TimerData, ByVal delayMillis As Long)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Function ITimerManager_tryTriggerTimer(ByVal data As TimerData) As Boolean
    Const loggerSourceName As String = "tryTriggerTimer"
    
    On Error GoTo catchError
    'Post fake message to queue to act as an already elapsed timer
    If WinAPI.PostMessage(this.messageWindow.handle, WM_TIMER, data.id, data.timerProc) = 0 Then
        throwDllError Err.LastDllError
    Else
        ITimerManager_tryTriggerTimer = True
    End If
    
cleanExit:
    Exit Function
    
catchError:
    If Err.Number = systemErrorCodes.ERROR_NOT_ENOUGH_QUOTA Then
        ITimerManager_tryTriggerTimer = False
        Resume cleanExit
        
    Else
        logError "WinAPI.PostMessage", Err.Number, Err.Description
        raiseInternalError APIPostMessageError, loggerSourceName
        Resume                                   'for debugging - break above and jump to the error-raising statement
        
    End If
End Function

Private Sub ITimerManager_KillAllTimers(ByRef dataArray() As TimerData)
    Const loggerSourceName As String = "KillAllTimers"
    
    'NOTE this procedure is called when raising errors so must not raise any itself
    On Error Resume Next
    log InfoLevel, loggerSourceName, printf("{0} registered timer(s)", UBound(dataArray) - LBound(dataArray)) 'TODO move this elswhere
    
    Set this.messageWindow = Nothing 'terminateMessageWindow - it's autoinstantiated so no tests

    If Err.Number <> 0 Then logError loggerSourceName, Err.Number, Err.Description
    On Error GoTo 0
End Sub


'@Description("Posts a fake message to the message queue - timerProc should be 0 for managed timers")
Private Function tryScheduleProc(ByVal windowHandle As LongPtr, ByVal timerID As LongPtr, ByVal timerProc As LongPtr) As Boolean

End Function

Private Sub KillTimer(ByVal windowHandle As LongPtr, ByVal timerID As LongPtr)
       
End Sub

'@Description("Terminates message window and removes stored timer references")
Private Sub KillAllTimers()

End Sub

Private Sub raiseInternalError(ByVal errorCode As InternalTimerError, Optional ByVal Source As String = "raiseInternalError")
      
    Select Case errorCode
        Case InternalTimerError.CreateMessageWindowError
            Err.Description = "Unable to obtain message window"
        
        Case InternalTimerError.APIKillTimerError
            Err.Description = "Error when calling API to destroy timer"
        
        Case InternalTimerError.APIStartTimerError
            Err.Description = "Error when calling API to create timer"
        
        Case InternalTimerError.CastKeyToWrapperError
            Err.Description = "Failed to cast key object to expected interface"
            
        Case InternalTimerError.APIPostMessageError
            Err.Description = "Failed to manually post a message to the queue"
            
        Case Else
            'rethrow error
            On Error Resume Next
            Err.Raise errorCode                  'fake raise to grab text for logging
            Dim errDescription As String
            errDescription = Err.Description
            On Error GoTo 0
            Err.Description = errDescription
            
    End Select
    'NOTE only log external errors as you can't rely on external loggers
    Err.Raise errorCode, Source

End Sub

'For testing
Friend Property Get messageWindowHandle() As LongPtr
    messageWindowHandle = this.messageWindow.handle
End Property
