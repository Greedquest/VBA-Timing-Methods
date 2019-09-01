VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TickerAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "API for setting up timers to callback functions, wraps Windows Timers"
'@Exposed
'@Folder("FirstLevelAPI")
'@PredeclaredID
'@ModuleDescription("API for setting up timers to callback functions, wraps Windows Timers")
Option Explicit

Public Enum TimerError
    CreateTimerError = vbObjectError + 1
    DestroyTimerError
    TimerNotFoundError
    SynchronousCallError
    TimerApiLockedError
    GetMessageWindowError
    [_End]
End Enum

Private Enum InternalTimerError
    [_Start] = TimerError.[_End] 'just in case of clashes, let's offset the errors
    CreateMessageWindowError
    APIKillTimerError
    CastKeyToWrapperError
    APIStartTimerError
    APIPostMessageError
End Enum

'Private Const Default_Lock As Boolean = True
Private Const Default_Max_Timer_Count As Long = 100
'Private Const Default_Message_Window_Suffix As String = "MessageWindow"


Private Type tCallback
    
    timerIDs As New Scripting.Dictionary              '{timerID: keyObj}
    'functionSet As New Dictionary
'    recursionStack As New Stack
'    isLocked As Boolean
    maxTimerCount As Long
    messageWindow As ModelessMessageWindow
End Type

Private this As tCallback

'@Description("Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error")
Public Function StartUnmanagedTimer(ByVal timerFunction As LongPtr, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
Attribute StartUnmanagedTimer.VB_Description = "Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error"
    Const loggerSourceName As String = "StartUnmanagedTimer"
    
    On Error GoTo getHandleFail
    Dim windowHandle As LongPtr
    windowHandle = getHandle
        
    On Error GoTo createTimerFail
    Dim newTimerID As LongPtr
    newTimerID = createTimer(windowHandle, UnmanagedCallbackWrapper.Create(timerFunction, data), delayMillis, timerFunction)
    StartUnmanagedTimer = newTimerID
    
    On Error GoTo scheduleProcFail
    If runImmediately Then
        'schedule immediate proc; requires callback for unmanaged timer
        If Not tryScheduleProc(windowHandle, newTimerID, timerFunction) Then
            'queue is too full right now, no point scheduling as it wouldn't be evaluated in time anyway
            'could try flushing the queue instead
            log WarnLevel, loggerSourceName, "Message queue is too full to post to, so cannot runImmediately"
        End If
    End If
    
    log InfoLevel, loggerSourceName, printf("UnmanagedTimer with id {0} created", newTimerID)
    Exit Function
    
getHandleFail:
    logError "getHandle", Err.Number, Err.Description
    raiseError GetMessageWindowError, loggerSourceName
    Resume 'for debugging - break above and jump to the error-raising statement
    
createTimerFail:
    logError "createTimer", Err.Number, Err.Description
    raiseError CreateTimerError, loggerSourceName
    Resume
    
scheduleProcFail:
    logError "scheduleProc", Err.Number, Err.Description
    KillTimerByID newTimerID 'NOTE may raise an error
    raiseError SynchronousCallError, loggerSourceName
    Resume
    
End Function

Public Function StartManagedTimer(ByVal timerFunction As ITimerProc, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
    Const loggerSourceName As String = "StartManagedTimer"
    
    On Error GoTo getHandleFail
    Dim windowHandle As LongPtr
    windowHandle = getHandle
        
    On Error GoTo createTimerFail
    Dim newTimerID As LongPtr
    newTimerID = createTimer(windowHandle, ManagedCallbackWrapper.Create(timerFunction, data), delayMillis)
    StartManagedTimer = newTimerID
    
    On Error GoTo scheduleProcFail
    If runImmediately Then
        'schedule immediate proc - No callback proc specified for managed timers
        If Not tryScheduleProc(windowHandle, newTimerID) Then
            'queue is too full right now, no point scheduling as it wouldn't be evaluated in time anyway
            'could try flushing the queue instead
            log WarnLevel, loggerSourceName, "Message queue is too full to post to, so cannot runImmediately"
        End If
    End If
    
    log InfoLevel, loggerSourceName, printf("ManagedTimer with id {0} created", newTimerID)
    Exit Function
    
getHandleFail:
    logError "getHandle", Err.Number, Err.Description
    raiseError GetMessageWindowError, loggerSourceName
    Resume 'for debugging - break above and jump to the error-raising statement
    
createTimerFail:
    logError "createTimer", Err.Number, Err.Description
    raiseError CreateTimerError, loggerSourceName
    Resume
    
scheduleProcFail:
    logError "scheduleProc", Err.Number, Err.Description
    KillTimerByID newTimerID 'NOTE may raise an error
    raiseError SynchronousCallError, loggerSourceName
    Resume
    
End Function

'@Description("API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimerByID(ByVal timerID As LongPtr)
Attribute KillTimerByID.VB_Description = "API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError"
    Const loggerSourceName As String = "KillTimerByID"
    
    If timerExists(timerID) Then
    
        On Error GoTo getHandleFail
        Dim windowHandle As LongPtr
        windowHandle = getHandle
        
        On Error GoTo killTimerFail
        killTimer windowHandle, timerID
            
        log InfoLevel, loggerSourceName, printf("Timer with id {0} destroyed", timerID)

    Else
    
        raiseError TimerNotFoundError, loggerSourceName
    End If
    
    Exit Sub
    
getHandleFail:
    logError "getHandle", Err.Number, Err.Description
    raiseError GetMessageWindowError, loggerSourceName
    Resume 'for debugging - break above and jump to the error-raising statement

killTimerFail:
    logError "killTimer", Err.Number, Err.Description
    raiseError DestroyTimerError, loggerSourceName
    Resume 'for debugging - break above and jump to the error-raising statement
    
End Sub

'@Description("API kills windows timers registered to a certain callback. No error is raised if none are found,failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimersByFunction(ByVal timerFunction As LongPtr)
Attribute KillTimersByFunction.VB_Description = "API kills windows timers registered to a certain callback. No error is raised if none are found,failure to destroy a registered ID raises DestroyTimerError"
    Const loggerSourceName As String = "KillTimersByFunction"
    
    Dim timerID As Variant 'key should be timerID
    Dim timerCount As Long
    For Each timerID In this.timerIDs.Keys
        If TypeOf this.timerIDs.Item(timerID) Is UnmanagedCallbackWrapper Then
            Dim wrapper As ICallbackWrapper
            Set wrapper = this.timerIDs.Item(timerID)
            If wrapper.Callback = timerFunction Then
                KillTimerByID timerID 'just bubble errors
                timerCount = timerCount + 1
            End If
        End If
    Next timerID
    log InfoLevel, loggerSourceName, printf("{0} timer{1} with function id {2} {3}", timerCount, IIf(timerCount = 1, vbNullString, "s"), timerFunction, IIf(timerCount = 0, "found", "destroyed"))
End Sub

'@Description("Terminates message window and removes stored timer references")
Private Sub killAllTimers()
Attribute killAllTimers.VB_Description = "Sub loops through all known timers and tries to kill them. If none exist, it tries killing them by ID until none are left. Should not raise errors. Clients should just re-initialise API"
    Const loggerSourceName As String = "killAllTimers"
    
    'NOTE this procedure is called when raising errors so must not raise any itself
    On Error Resume Next
    log InfoLevel, loggerSourceName, printf("{0} registered timer(s)", this.timerIDs.Count)

    this.timerIDs.RemoveAll
    terminateMessageWindow

    If Err.Number <> 0 Then logError loggerSourceName, Err.Number, Err.Description
End Sub

Private Sub clearMessages(ByVal timerID As LongPtr)
    Dim message As WinAPI.timerMessage
    Dim i As Long
    'TODO Will there be messages here? If scheduleProc followed by killTimer then yes - but those are managed messages which we can filter by id so no probs
    If GetQueueStatus(QS_TIMER).HiWord <> 0 Then
        log WarnLevel, "clearMessages", "Messages in this queue might be important:)"
'        Do While WinAPI.PeekTimerMessage(message, this.messageWindow.handle, WM_TIMER, WM_TIMER, PM_REMOVE) <> 0 And i < this.maxTimerCount
'            'TODO if timerids specified then we can re-send accidentally removed messages - ones which don't match ids
'            i = i + 1
'            If message.timerID <> timerID Then WinAPI.DispatchTimerMessage message
'        Loop
    End If
    log InfoLevel, "clearMessages", printf("Cleared {0} messages", i)
End Sub

Friend Property Get timerExists(ByVal key As LongPtr) As Boolean
    On Error Resume Next 'if there's a problem then the timerID is as good as unregistered anyway
    timerExists = this.timerIDs.Exists(key)
    On Error GoTo 0
End Property

'@Description("Posts a fake message to the message queue - timerProc should be 0 for managed timers")
Private Function tryScheduleProc(ByVal windowHandle As LongPtr, ByVal timerID As LongPtr, Optional ByVal timerProc As LongPtr = 0) As Boolean
    Const loggerSourceName As String = "tryScheduleProc"
    
    On Error GoTo catchError
    'Post fake message to queue to act as an already elapsed timer
    If WinAPI.PostMessage(windowHandle, WM_TIMER, timerID, timerProc) = 0 Then
        throwDllError Err.LastDllError
    Else
        tryScheduleProc = True
    End If
    
cleanExit:
    Exit Function
    
catchError:
    If Err.Number = systemErrorCodes.ERROR_NOT_ENOUGH_QUOTA Then
        tryScheduleProc = False
        Resume cleanExit
        
    Else
        logError "killTimer", Err.Number, Err.Description
        raiseInternalError APIPostMessageError, loggerSourceName
        Resume 'for debugging - break above and jump to the error-raising statement
        
    End If
End Function

'@Description("Creates a timer with specified delay and returns its ID. Takes a uniqie object which wraps data to pass to timers and also to use as a key")
Private Function createTimer(ByVal windowHandle As LongPtr, ByVal key As Object, ByVal delayMillis As Long, Optional ByVal timerProc As LongPtr = 0) As LongPtr
    Const loggerSourceName As String = "createTimer"
    
'    On Error GoTo castWrapperFail
'    'pass key as object since it's easy to downcast to ICallbackWrapper interface, but hard to up cast from ICallbackWrapper to every possible implementation
'    Dim wrapper As ICallbackWrapper
'    Set wrapper = key
'    'key is either the TickerParams or the TimerProcWrapper so could upcast each
'
'    'Create timer with ID of the key object; guaranteed to be unique
'    'passing 0 as timerFunction will result in message posted to messageWindow - managed timer
'

    'Just bubble type mismatch errors as there's no cleanup to do
    If timerProc = 0 Then
        If TypeOf key Is ManagedCallbackWrapper Then
            'managed timers use the default TIMERPROC
            timerProc = VBA.CLngPtr(AddressOf ManagedTimerCallbackInvoker)
        Else
            'all others should specify one (i.e it shouldn't be zero)
            Err.Raise 5, Description:="Only managed timers may set timerProc to zero"
        End If
    End If
    
    On Error GoTo setTimerFail 'Custom handler so we can log precise dll errors and condense error messages + clear up any timer which may have been made
    
    Dim newTimerID As LongPtr
    newTimerID = WinAPI.SetTimer(windowHandle, ObjPtr(key), delayMillis, timerProc)
    
    If newTimerID = 0 Then
        throwDllError Err.LastDllError
    
    ElseIf newTimerID <> ObjPtr(key) Then
        Err.Raise 5, Description:="timerID does not have expected value"
        
    Else
        'store copy of key object so it doesn't go out of scope
        Set this.timerIDs.Item(newTimerID) = key
        
    End If
    
    createTimer = newTimerID
    Exit Function
    
'castWrapperFail:
'    logError "castWrapper", Err.Number, Err.Description
'    raiseInternalError CastKeyToWrapperError, loggerSourceName
'    Resume 'for debugging - break above and jump to the error-raising statement

setTimerFail:
    logError "WinAPI.SetTimer", Err.Number, Err.Description
    killTimer windowHandle, newTimerID
    raiseInternalError CastKeyToWrapperError, loggerSourceName
    Resume 'for debugging - break above and jump to the error-raising statement
    
    
End Function

Private Sub killTimer(ByVal windowHandle As LongPtr, ByVal timerID As LongPtr)
    'TODO: could clear messages for this timer

    On Error GoTo cleanFail
    
    '0 indicates some failure
    If WinAPI.killTimer(windowHandle, timerID) = 0 Then
        throwDllError Err.LastDllError, "Call returned zero, probably tried to kill non-existent timer"
    End If
    
cleanExit:
    If timerExists(timerID) Then this.timerIDs.Remove timerID             'remove regardless of success, since it's either gone now, or was invalid anyway
    Exit Sub
    
cleanFail:
    logError "WinAPI.killTimer", Err.Number, Err.Description
    raiseInternalError APIKillTimerError, "killTimer"
    Resume cleanExit
    
End Sub

Private Sub raiseError(ByVal errorCode As TimerError, Optional ByVal Source As String = "raiseError")

    log TraceLevel, "raiseError", "Destroying timers and unsubclassing window so error can be raised"
    killAllTimers 'cleanly un subclasses message window as this prevents errors being raised
    
    Select Case errorCode
        Case TimerError.DestroyTimerError
            Err.Description = "Uh Oh, can't kill the timer :("
            
        Case TimerError.TimerNotFoundError
            Err.Description = "Timer not found"
            
        Case TimerError.CreateTimerError
            Err.Description = "Couldn't create Timer"
            
        Case TimerError.SynchronousCallError
            Err.Description = "Error when running synchronously"
            
        Case TimerError.GetMessageWindowError
            Err.Description = "Couldn't obtain a handle to the message window"
            
        Case Else
            'rethrow error
            On Error Resume Next
                Err.Raise errorCode 'fake raise to grab text for logging
                Dim errDescription As String
                errDescription = Err.Description
            On Error GoTo 0
            Err.Description = errDescription
            
    End Select
    
    logError Source, errorCode, Err.Description
    
    Err.Raise errorCode


    
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
                Err.Raise errorCode 'fake raise to grab text for logging
                Dim errDescription As String
                errDescription = Err.Description
            On Error GoTo 0
            Err.Description = errDescription
            
    End Select
    'logError Source, errorCode, Err.Description
    Err.Raise errorCode

End Sub

Private Function getMessageWindow() As ModelessMessageWindow
    If this.messageWindow Is Nothing Then
        Set this.messageWindow = New ModelessMessageWindow
'        'make a new one
'        If ModelessMessageWindow.tryCreate(this.messageWindow, AddressOf MessageWindowProcs.ManagedTimerMessageWindowSubclassProc) Then
'            MessageWindowProcs.Activate 'otherwise WM_TIMER messages are ignored by the window
'            Set getMessageWindow = this.messageWindow
'        Else
'            raiseInternalError CreateMessageWindowError, "getMessageWindow"
'        End If
    End If
    Set getMessageWindow = this.messageWindow
End Function

Private Function getHandle() As LongPtr
    getHandle = getMessageWindow.handle 'will bubble up any errors
End Function

'Private Sub setupMessageWindow()
'    If ModelessMessageWindow.tryCreate(this.messageWindow, AddressOf MessageWindowProcs.ManagedTimerMessageWindowSubclassProc) Then
'        MessageWindowProcs.Activate 'otherwise WM_TIMER messages are ignored by the window
'    Else
'        log FatalLevel, "setupMessageWindow", "Unable to hook message window"
'    End If
'End Sub

Private Sub terminateMessageWindow()
    'NOTE following an error, messageWindow is terminated but not reset
    If this.messageWindow Is Nothing Then
        log WarnLevel, "terminateMessageWindow", "Message window is not set"
    Else
        log DebugLevel, "terminateMessageWindow", "Destroying message window" '"Removing SUBCLASSPROCs: " & IIf(this.messageWindow.tryRemoveAllSubclasses, "successful", "unsuccessful")
        Set this.messageWindow = Nothing
    End If
End Sub

Friend Function tryGetManagedWrapper(ByVal timerID As LongPtr, ByRef outWrapper As ManagedCallbackWrapper) As Boolean
    'Pretty much the only valid use of checkedException handling
    If this.timerIDs.Exists(timerID) Then
        On Error Resume Next
        Set outWrapper = this.timerIDs.Item(timerID)
        tryGetManagedWrapper = Err.Number = 0
        On Error GoTo 0
    Else
        tryGetManagedWrapper = False
    End If
End Function

Private Sub Class_Initialize()
    'Set up defaults
    this.maxTimerCount = Default_Max_Timer_Count
    'setupMessageWindow
End Sub

Private Sub Class_Terminate()
    'To free up this resource we need to unsubclass the window. It will already be free if an error was just raised
    If Not this.messageWindow Is Nothing Then terminateMessageWindow
End Sub

'For Testing
Friend Property Get messageWindowHandle() As LongPtr
    messageWindowHandle = getHandle
End Property

Friend Sub makeBadness()
    raiseError 5
End Sub
