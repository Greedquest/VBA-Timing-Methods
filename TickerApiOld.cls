VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TickerApiOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "API for setting up timers to callback functions, wraps Windows Timers"
'@Exposed
'@Folder("Old.FirstLevelAPI")
'@PredeclaredID
'@ModuleDescription("API for setting up timers to callback functions, wraps Windows Timers")
'@IgnoreModule
Option Explicit
'
'Public Enum TimerError
'    CreateTimerError = vbObjectError + 1
'    DestroyTimerError
'    TimerNotFoundError
'    SynchronousCallError
'    TimerApiLockedError
'End Enum

Private Const Default_Lock As Boolean = True
Private Const Default_Max_Timer_Count As Long = 100
Private Const Default_Message_Window_Suffix As String = "MessageWindow"
                           
Private Type tWindow
    handle As LongPtr
    name As String
End Type

Private Type tCallback
    timerIDs As Scripting.Dictionary             'timerID: keyObj
    'functionSet As New Dictionary
    recursionStack As New Stack
    isLocked As Boolean
    maxTimerCount As Long
    messageWindow As tWindow
End Type

Private this As tCallback

'@Description("Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error")
Public Function StartUnmanagedTimer(ByVal timerFunction As LongPtr, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
Attribute StartUnmanagedTimer.VB_Description = "Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error"
    checkIfLocked "StartUnmanagedTimer"
    
    '        this.recursionStack.Push makeTimer
    '        If Not tryCallProc(timerFunction, makeTimer) Then 'NOTE this will always be true as errors are untrappable and pointer is assumed valid
    '            raiseError SynchronousCallError, "StartTimer"
    '        End If
    '        'Check whether timerFunction has edited the flag on the top of the stack - could do a 1 liner
    '        Set makeTimer = this.recursionStack.Pop
    '        If Not makeTimer.TickerIsRunning Then Exit Function
    
    
    Dim newTimerID As LongPtr
    If Not tryCreateTimer(UnmanagedCallbackWrapper.Create(timerFunction, data), delayMillis, newTimerID) Then
        raiseError CreateTimerError, "StartUnmanagedTimer"
    Else
        StartUnmanagedTimer = newTimerID
    End If
    If runImmediately Then
        'schedule immediate proc - No callback proc specified
        If Not tryScheduleProc(newTimerID, timerFunction) Then
            KillTimerByID newTimerID             'NOTE may raise error
        End If
    End If
End Function

Public Function StartManagedTimer(ByVal timerFunction As ITimerProc, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
    checkIfLocked "StartManagedTimer"
    Dim newTimerID As LongPtr
    If Not tryCreateTimer(ManagedCallbackWrapper.Create(timerFunction, data), delayMillis, newTimerID) Then
        raiseError CreateTimerError, "StartManagedTimer"
    Else
        StartManagedTimer = newTimerID
    End If
    If runImmediately Then
        'schedule immediate proc - No callback proc specified
        If Not tryScheduleProc(newTimerID) Then
            KillTimerByID newTimerID             'NOTE may raise error
        End If
    End If
End Function

'@Description("API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimerByID(ByVal timerID As LongPtr)
Attribute KillTimerByID.VB_Description = "API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError"
    checkIfLocked "KillTimerByID"
    If timerExists(timerID) Then
        If Not tryKillTimer(timerID) Then
            raiseError DestroyTimerError, "KillTimerByID"
        Else
            log InfoLevel, "KillTimerByID", printf("Timer with id {0} destroyed", timerID)
        End If
    Else
        raiseError TimerNotFoundError, "KillTimerByID"
    End If
End Sub

'@Description("API kills windows timers registered to a certain callback. No error is raised if none are found,failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimersByFunction(ByVal timerFunction As LongPtr)
Attribute KillTimersByFunction.VB_Description = "API kills windows timers registered to a certain callback. No error is raised if none are found,failure to destroy a registered ID raises DestroyTimerError"
    checkIfLocked "KillTimersByFunction"
    Dim key As Variant 'key should be timerID
    Dim timerCount As Long
    For Each key In this.timerIDs.Keys
        If TypeOf this.timerIDs.Item(key) Is UnmanagedCallbackWrapper Then
            Dim wrapper As ICallbackWrapper
            Set wrapper = this.timerIDs.Item(key)
            If wrapper.Callback = timerFunction Then
                If Not tryKillTimer(key) Then
                    raiseError DestroyTimerError, "KillTimersByFunction"
                    Exit For
                End If
                timerCount = timerCount + 1
            End If
        End If
    Next key
    log InfoLevel, "KillTimersByFunction", printf("{0} timer{1} with function id {2} {3}", timerCount, IIf(timerCount = 1, vbNullString, "s"), timerFunction, IIf(timerCount = 0, "found", "destroyed"))
End Sub

'@Description("Sub loops through all known timers and tries to kill them. If none exist, it tries killing them by ID until none are left. Should not raise errors. Clients should just re-initialise API")
Friend Sub killAllTimers()
Attribute killAllTimers.VB_Description = "Sub loops through all known timers and tries to kill them. If none exist, it tries killing them by ID until none are left. Should not raise errors. Clients should just re-initialise API"
    'NOTE this procedure is called when raising errors so must not raise any itself
        
    log DebugLevel, "KillAllTimers", printf("{0} registered timer(s)", this.timerIDs.Count)
    
    'Try to just remove message window, if that fails then kill manually
    If this.timerIDs.Count = 0 Then
        'do nothing
    ElseIf tryDestroyMessageWindowByName(this.messageWindow.name, this.messageWindow.handle) Then
        'NOTE destroying message window should orphan all the existing timers and reset the message queue, so no need for further clear up
'        For Each timerID In this.timerIDs
'            tryKillTimer timerID
'        Next timerID
'        clearAllMessages this.timerIDs
        this.timerIDs.RemoveAll
        setupMessageWindow 'TODO should this be here, probably not
    Else
        Dim timerID As Variant
        Dim destroyedTimers As New Collection
        For Each timerID In this.timerIDs.Keys
            On Error Resume Next                 'no point returning success since there's nothing we can do about it
            KillTimerByID timerID
            destroyedTimers.Add timerID
            On Error GoTo 0
            log DebugLevel, "KillAllTimers", printf("Killing timer #{0}", timerID)
        Next timerID
        clearAllMessages destroyedTimers
        'this.timerIDs.RemoveAll
    End If
           
End Sub

Private Sub clearAllMessages(Optional ByVal timerIDs As Dictionary)
    'TODO when custom window handle is used, there will be no need to validate timerIDs
    Dim tmpMsg As tagMSG
    Dim i As Long
    Do While WinAPI.PeekMessage(tmpMsg, this.messageWindow.handle, WM_TIMER, WM_TIMER, PM_REMOVE) <> 0 And i < this.maxTimerCount
        'TODO if timerids specified then we can re-send accidentally removed messages - ones which don't match ids
        i = i + 1
    Loop
    log InfoLevel, "clearAllMessages", printf("Cleared {0} messages", i)
End Sub

Friend Property Get timerExists(ByVal key As LongPtr) As Boolean
    timerExists = this.timerIDs.Exists(key)
End Property

'Private Function tryCallProc(ByVal timerFunction As LongPtr, ByVal timerFlag As TickerParams) As Boolean
'    On Error GoTo cleanFail
'    CallWindowProc timerFunction, timerFlag, WM_NOTIFY 'not sure what the return value of this is so no useful try result
'    tryCallProc = True
'
'cleanExit:
'    Exit Function
'
'cleanFail:
'    logError "tryCallProc", Err.Number, Err.Description
'    tryCallProc = False
'    Resume cleanExit
'
'End Function

'@Description("Posts a fake message to the message queue - timerProc should be 0 for managed timers")
Private Function tryScheduleProc(ByVal timerID As LongPtr, Optional ByVal timerProc As LongPtr = 0) As Boolean
Attribute tryScheduleProc.VB_Description = "Posts a fake message to the message queue - timerProc should be 0 for managed timers"
    
    'Post fake message to queue
    tryScheduleProc = WinAPI.PostMessage(this.messageWindow.handle, WM_TIMER, timerID, timerProc) <> 0
        
End Function

'@Description("Creates a timer with specified delay and returns its ID. Takes an object to use as key, ")
Private Function tryCreateTimer(ByVal key As Object, ByVal delayMillis As Long, ByRef outTimerID As LongPtr) As Boolean
Attribute tryCreateTimer.VB_Description = "Creates a timer with specified delay and returns its ID. Takes an object to use as key, "
      
    On Error GoTo cleanFail
    
    'pass key as object since it's easy to downcast to ICallbackWrapper interface, but hard to up cast from ICallbackWrapper to every possible implementation
    Dim wrapper As ICallbackWrapper
    Set wrapper = key
    'key is either the TickerParams or the TimerProcWrapper so could upcast each
    
    'Create timer with ID of the key object; guaranteed to be unique
    'If timerExists(ObjPtr(key)) Then          'only make a new timer if there isn't one already
    'passing 0 as timerFunction will result in message posted to messageWindow - managed timer
    
    outTimerID = WinAPI.SetTimer(this.messageWindow.handle, ObjPtr(key), delayMillis, IIf(wrapper.CallbackType = cbRawPointer, wrapper.Callback, 0))
    If outTimerID = ObjPtr(key) Then
        Set this.timerIDs.Item(outTimerID) = key 'store copy of key object so it doesn't go out of scope
        tryCreateTimer = True
    Else
        tryCreateTimer = False
    End If
    'End If
    
cleanExit:
    Exit Function
    
cleanFail:
    logError "tryCreateTimer", Err.Number, Err.Description
    tryCreateTimer = False
    Resume cleanExit
    
End Function

Private Function tryKillTimer(ByVal timerID As Long) As Boolean
    'TODO: could clear messages for this timer
    On Error GoTo cleanFail
    If timerExists(timerID) Then
        tryKillTimer = WinAPI.KillTimer(this.messageWindow.handle, timerID) <> 0 'kill the timer associated with that fn  & ID
        this.timerIDs.Remove timerID             'remove regardless of success, since it's either gone now, or was invalid anyway
    Else
        tryKillTimer = WinAPI.KillTimer(this.messageWindow.handle, timerID) <> 0
    End If

cleanExit:
    Exit Function
    
cleanFail:
    logError "tryKillTimer", Err.Number, Err.Description
    tryKillTimer = False
    Resume cleanExit
    
End Function

Private Function getUniqueID() As Long
    'TODO if we keep track of ids then they should always be consecutive with max 1 gap
    Static i As Long
    i = i + 1
    getUniqueID = i
    log DebugLevel, "getUniqueID", printf("ID {0} generated", i)
End Function

Private Sub log(ByVal loggerLevel As LogLevel, ByVal Source As String, ByVal message As String)
    If Not LogManager.IsEnabled(loggerLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("TickerAPI" & loggerLevel, loggerLevel)
    End If
    LogManager.log loggerLevel, Toolbox.Strings.Format("{0} - {1}", Source, message)
End Sub

Private Sub logError(ByVal Source As String, ByVal errNum As Long, ByVal errDescription As String)
    If Not LogManager.IsEnabled(ErrorLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("TickerAPI", ErrorLevel)
    End If
    LogManager.log ErrorLevel, Toolbox.Strings.Format("{0} raised an error: #{1} - {2}", Source, errNum, errDescription)
End Sub

Private Sub raiseError(ByVal errorCode As TimerError, Optional ByVal Source As String = "raiseError")
    'BUG raising an unhandled error crashes if there are any messages queued, so we also need to flush the queue before raising errors
    killAllTimers
    Select Case errorCode
        Case TimerError.DestroyTimerError
            Err.Description = "Uh Oh, can't kill the timer :("
        Case TimerError.TimerNotFoundError
            Err.Description = "Timer not found"
        Case TimerError.CreateTimerError
            Err.Description = "Couldn't create Timer"
        Case TimerError.SynchronousCallError
            Err.Description = "Error when running synchronously"
        Case TimerError.TimerApiLockedError
            Err.Description = "Call the Unlock method before attempting to use the API"
        Case Else
            'rethrow error
    End Select
    logError Source, errorCode, Err.Description
    Err.Raise errorCode
End Sub

Private Sub checkIfLocked(Optional ByVal Source As String = "checkIfLocked")
        If this.isLocked Then
            If this.timerIDs.Count = 0 Then          'message didn't come from some timer hanging around, so we can unlock safely
                'NOTE an unlock timer will only be made if there are other timers hanging around, so we don't need to check for it here
                log InfoLevel, "checkIfLocked", "No timers hanging around, so auto-unlock run"
                UnlockApi
            Else                                     'there is a timer hanging around when this was called, so it *could* be responsible, so better safe than sorry
                raiseError TimerApiLockedError, Source
            End If
        End If
End Sub

'@Description("Remove restrictions which prevent certain methods from being called by orphan procs; should be placed in all user code but never in code run by TIMERPROCs")
Public Sub UnlockApi()
    'TODO permanent lock following fatal error
    log DebugLevel, "UnlockApi", "API unlocking called" & IIf(this.isLocked, vbNullString, " (already unlocked)")
    this.isLocked = False
End Sub

'@Description("Helper function for un-managed timers, should be called at least once following a state loss by any timer-procs that may be running")
Public Sub Poke()
    'no-op just triggers Class_Initialize
End Sub

Public Function tryScheduleUnlock() As Boolean

    'messages must be associated with a valid timer, so let's make one
    'we have to do it manualy as StartUnmanagedTimer checks lock status
    Const INFINITE_INTERVAL As Long = &H7FFFFFFF
    Dim validationTimerID As LongPtr
    If tryCreateTimer(UnmanagedCallbackWrapper.Create(AddressOf TickerAPITimerProcs.UnlockCallbackProc), INFINITE_INTERVAL, validationTimerID) Then
        If tryScheduleProc(validationTimerID, AddressOf TickerAPITimerProcs.UnlockCallbackProc) Then
            tryScheduleUnlock = True
        Else
            If Not tryKillTimer(validationTimerID) Then killAllTimers 'don't want to leave one hanging around
        End If
    End If
    
End Function

Private Function setupMessageWindow() As Boolean
    'This should work whether or not a window & handle already exist
    If tryHookMessageHandler(AddressOf MessageWindowProcs.ManagedTimerMessageWindowSubclassProc, this.messageWindow.name, this.messageWindow.handle) Then
        'TODO this message should reflectg whether a new one was made
        log DebugLevel, "HookMessages", printf("Reset message loop for window '{0}'", this.messageWindow.name)
        setupMessageWindow = True
    Else
        log FatalLevel, "HookMessages", printf("tryHookMessageHandler can't find or set up '{0}' message window to handle timers", this.messageWindow.name)
    End If
End Function

Private Sub Class_Initialize()
    'Set up defaults
    this.isLocked = Default_Lock
    this.maxTimerCount = Default_Max_Timer_Count
    this.messageWindow.name = TypeName(Me) & Default_Message_Window_Suffix
    
    'Kill existing timers
    Set this.timerIDs = cache.loadObject(TypeName(Me) & ".timerIDs", New Scripting.Dictionary)
    killAllTimers
    If setupMessageWindow Then
        'this should trigger unlock after the current synchronous call - whichever one triggered the predeclared class to initialise
        
        If this.isLocked Then
            If this.timerIDs.Count = 0 Then
                'we know class wasn't started by an orphaned timer, so no need for a lock - unlock immediately
                log InfoLevel, "Class_Initialize", "No timers hanging around, so auto-unlock run immediately"
                UnlockApi
            Else
                'class may have been started by user or orphaned timer, so best to schedule an unlock to run after whatever code called us
                If Not tryScheduleUnlock Then log ErrorLevel, "tryScheduleUnlock", "Unable to post auto-unlock message, if this problem persists consider a managed timer"
            End If
        End If
    Else
        'Fatal error, don't let anyone use the class.
        log FatalLevel, "Class_Initialize", "Couldn't get message window"
        this.isLocked = True
    End If
End Sub

Private Sub Class_Terminate()
    killAllTimers
    tryDestroyMessageWindowByName (this.messageWindow.name)
End Sub

'For Testing
Friend Property Get messageWindowHandle() As LongPtr
    messageWindowHandle = this.messageWindow.handle
End Property

