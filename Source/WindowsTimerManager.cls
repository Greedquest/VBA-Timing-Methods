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
    On Error GoTo CleanFail
    
    '0 indicates some failure
    If WinAPI.KillTimer(this.messageWindow.handle, data.ID) = 0 Then
        throwDllError Err.LastDllError, "Call returned zero, probably tried to kill non-existent timer"
    End If
    
CleanExit:
    Exit Sub
    
CleanFail:
    logError "WinAPI.KillTimer", Err.Number, Err.Description
    raiseInternalError APIKillTimerError, "KillTimer"
    Resume CleanExit
    
End Sub

Private Sub ITimerManager_StartTimer(ByVal data As TimerData, ByVal delayMillis As Long)
    Const loggerSourceName As String = "StartTimer"
    
    'Custom handler so we can log precise dll errors and condense error messages + clear up any timer which may have been made
    On Error GoTo setTimerFail
    
    Dim newTimerID As LongPtr
    newTimerID = WinAPI.SetTimer(this.messageWindow.handle, data.callbackWrapperPointer, delayMillis, data.timerProc)
    
    If newTimerID = 0 Then
        throwDllError Err.LastDllError
    
    ElseIf newTimerID <> data.ID Then
        Err.Raise 5, Description:="timerID does not have expected value" 'REVIEW is there a better assertion error to raise?
        
    End If
    
    Exit Sub
    
setTimerFail:
    logError "WinAPI.SetTimer", Err.Number, Err.Description
    ITimerManager_KillTimer data
    raiseInternalError APIStartTimerError, loggerSourceName
    Resume                                       'for debugging - break above and jump to the error-raising statement

End Sub

'TODO never used
Private Sub ITimerManager_UpdateTimer(ByVal data As TimerData, ByVal delayMillis As Long)
    'NOTE just an alias for windows timers, maybe not for others
    ITimerManager_StartTimer data, delayMillis
End Sub

Private Function ITimerManager_tryTriggerTimer(ByVal data As TimerData) As Boolean
    Const loggerSourceName As String = "tryTriggerTimer"
    
    On Error GoTo catchError
    'Post fake message to queue to act as an already elapsed timer
    If WinAPI.PostMessage(this.messageWindow.handle, WM_TIMER, data.ID, data.timerProc) = 0 Then
        throwDllError Err.LastDllError
    Else
        ITimerManager_tryTriggerTimer = True
    End If
    
CleanExit:
    Exit Function
    
catchError:
    If Err.Number = systemErrorCodes.ERROR_NOT_ENOUGH_QUOTA Then
        ITimerManager_tryTriggerTimer = False
        Resume CleanExit
        
    Else
        logError "WinAPI.PostMessage", Err.Number, Err.Description
        raiseInternalError APIPostMessageError, loggerSourceName
        Resume                                   'for debugging - break above and jump to the error-raising statement
        
    End If
End Function

Private Sub ITimerManager_KillAllTimers(ByVal dataArray As Variant)
    Const loggerSourceName As String = "KillAllTimers"
    
    'NOTE this procedure is called when raising errors so must not raise any itself
    On Error Resume Next
    log InfoLevel, loggerSourceName, printf("{0} registered timer(s)", UBound(dataArray) - LBound(dataArray)) 'TODO move this elswhere
    
    Set this.messageWindow = Nothing             'terminateMessageWindow - it's autoinstantiated so no tests

    If Err.Number <> 0 Then logError loggerSourceName, Err.Number, Err.Description
    On Error GoTo 0
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
            'Rethrow error
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
