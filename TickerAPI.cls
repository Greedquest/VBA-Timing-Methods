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
End Enum

Private Const Default_Lock As Boolean = True
Private Const Default_Max_Timer_Count As Long = 100
Private Const Default_Message_Window_Suffix As String = "MessageWindow"
                           
Private Type tWindow
    handle As LongPtr
    name As String
End Type

Private Type tCallback
    timerIDs As Scripting.Dictionary             'SetTimer ID: Callback Fn
    'functionSet As New Dictionary
    recursionStack As New Stack
    isLocked As Boolean
    maxTimerCount As Long
    messageWindow As tWindow
End Type

Private this As tCallback

'@Description("Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error")
Public Function StartTimer(ByVal timerFunction As LongPtr, Optional ByVal data As Variant, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500) As LongPtr
Attribute StartTimer.VB_Description = "Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error"
    checkIfLocked "StartTimer"
    Dim makeTimer As TickerParams
    Set makeTimer = TickerParams.Create(True, data)
    
    If runImmediately Then

'        this.recursionStack.Push makeTimer
'        If Not tryCallProc(timerFunction, makeTimer) Then 'NOTE this will always be true as errors are untrappable and pointer is assumed valid
'            raiseError SynchronousCallError, "StartTimer"
'        End If
'        'Check whether timerFunction has edited the flag on the top of the stack - could do a 1 liner
'        Set makeTimer = this.recursionStack.Pop
'        If Not makeTimer.TickerIsRunning Then Exit Function
    End If
    
    Dim newTimerID As LongPtr
    If Not tryCreateTimer(timerFunction, delayMillis, makeTimer, newTimerID) Then
        raiseError CreateTimerError, "StartTimer"
    Else
        StartTimer = newTimerID
    End If
End Function

Public Function StartManagedTimer(ByVal timerFunction As IFunction, Optional ByVal data As Variant, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500) As LongPtr
    checkIfLocked "StartManagedTimer"
    If runImmediately Then
'        Dim makeTimer As TickerParams
'        Set makeTimer = TickerParams.Create(True)
'        this.recursionStack.Push makeTimer
'        If Not tryCallProc(timerFunction, makeTimer) Then 'NOTE this will always be true as errors are untrappable and pointer is assumed valid
'            raiseError SynchronousCallError, "StartTimer"
'        End If
'        'Check whether timerFunction has edited the flag on the top of the stack - could do a 1 liner
'        Set makeTimer = this.recursionStack.Pop
'        If Not makeTimer.TickerIsRunning Then Exit Function
    End If
    Dim newTimerID As Long
    If Not tryCreateTimer(timerFunction, delayMillis, Nothing, newTimerID) Then
        raiseError CreateTimerError, "StartManagedTimer"
    Else
        StartManagedTimer = newTimerID
    End If
End Function

'@Description("API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimerByID(ByVal timerID As LongPtr)
Attribute KillTimerByID.VB_Description = "API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError"
    checkIfLocked "KillTimerByID"
    If timerExists(timerID) Then
        If Not tryKillTimer(timerID) Then
            raiseError DestroyTimerError, "KillTimerByID"
        End If
    Else
        raiseError TimerNotFoundError, "KillTimerByID"
    End If
End Sub

'@Description("API kills windows timers registered to a certain callback. No error is raised if none are found,failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimersByFunction(ByVal timerFunction As LongPtr)
Attribute KillTimersByFunction.VB_Description = "API kills windows timers registered to a certain callback. No error is raised if none are found,failure to destroy a registered ID raises DestroyTimerError"
    checkIfLocked "KillTimersByFunction"
    Dim key As Variant
    For Each key In this.timerIDs.Keys
        If this.timerIDs(key) = timerFunction Then
            If Not tryKillTimer(key) Then
                raiseError DestroyTimerError, "KillTimersByFunction"
                Exit For
            End If
        End If
    Next key
End Sub

'@Description("Sub loops through all known timers and tries to kill them. If none exist, it tries killing them by ID until none are left. Should not raise errors. Clients should just re-initialise API")
Friend Sub killAllTimers()
Attribute killAllTimers.VB_Description = "Sub loops through all known timers and tries to kill them. If none exist, it tries killing them by ID until none are left. Should not raise errors. Clients should just re-initialise API"
    'NOTE this procedure is called when raising errors so must not raise any itself
        
    log DebugLevel, "KillAllTimers", printf("{0} registered timer(s)", this.timerIDs.Count)
    
    'Try to just remove message window, if that fails then kill manually
    Dim timerID As Variant
    If this.timerIDs.Count = 0 Then
        'do nothing
    ElseIf tryDestroyMessageWindowByName(this.messageWindow.name, this.messageWindow.handle) Then
        'NOTE this stuff is just in case tryKillTimer does other cleaning up
        For Each timerID In this.timerIDs
            tryKillTimer timerID
        Next timerID
        clearAllMessages this.timerIDs
        this.timerIDs.RemoveAll
    Else
        For Each timerID In this.timerIDs
            On Error Resume Next                 'no point returning success since there's nothing we can do about it
            WinAPI.KillTimer 0, this.timerIDs(timerID)
            log DebugLevel, "KillAllTimers", printf("Killing timer #{0}", timerID)
        Next timerID
        clearAllMessages this.timerIDs
        this.timerIDs.RemoveAll
    End If
           
End Sub

Private Sub clearAllMessages(Optional ByVal timerIDs As Dictionary)
    'TODO when custom window handle is used, there will be no need to validate timerIDs
    Dim tmpMsg As tagMSG
    Dim i As Long
    Do While WinAPI.PeekMessage(tmpMsg, this.messageWindow.handle, WM_TIMER, WM_TIMER, PM_REMOVE) And i < this.maxTimerCount
        'TODO if timerids specified then we can re-send accidentally removed messages - ones which don't match ids
        i = i + 1
    Loop
    log InfoLevel, "clearAllMessages", printf("Cleared {0} messages", i)
End Sub

Friend Property Get timerExists(ByVal key As LongPtr) As Boolean
Attribute timerExists.VB_Description = "Checks the runningSet for the item with key timerFunction, returns the item if found, or false otherwise"
    timerExists = this.timerIDs.Exists(key)
End Property

Private Function tryCallProc(ByVal timerFunction As LongPtr, ByVal timerFlag As TickerParams) As Boolean
    On Error GoTo cleanFail
    CallWindowProc timerFunction, timerFlag, WM_NOTIFY 'not sure what the return value of this is so no useful try result
    tryCallProc = True
    
cleanExit:
    Exit Function
    
cleanFail:
    logError "tryCallProc", Err.Number, Err.Description
    tryCallProc = False
    Resume cleanExit
    
End Function

'@Description("Schedules something to run immediately by posting a message straight to the queue. Any timerProc should kill the timer created")
Private Function tryScheduleProc(ByVal timerProc As LongPtr) As Boolean
Attribute tryScheduleProc.VB_Description = "Schedules something to run immediately by posting a message straight to the queue. Any timerProc should kill the timer created"

    'messages must be associated with a valid timer, so let's make one
    Dim validationTimerID As Long
    Const INFINITE_INTERVAL As Long = &H7FFFFFFF
    If Not tryCreateTimer(timerProc, validationTimerID, INFINITE_INTERVAL) Then
        tryScheduleProc = False
        Exit Function
    End If
    
    'Post fake message to queue
    If PostMessage(this.messageWindow.handle, WM_TIMER, validationTimerID, timerProc) = False Then
        tryKillTimer validationTimerID
        tryScheduleProc = False
        Exit Function
    End If
    
    tryScheduleProc = True
        
End Function

Private Function tryCreateTimer(ByVal timerFunction As Variant, ByVal delayMillis As Long, ByVal params As TickerParams, ByRef outTimerID As LongPtr) As Boolean
    On Error GoTo cleanFail
    'key is either the TickerParams or the IFunction
    Dim key As Object
    Set key = IIf(IsObject(timerFunction), timerFunction, params)
    'Create timer with ID of the key object; guaranteed to be unique
    If key Is Not Nothing And Not timerExists(key) Then          'only make a new timer if there isn't one already
        'passing 0 as timerFunction will result in message posted to messageWindow - managed timer
        outTimerID = WinAPI.SetTimer(this.messageWindow.handle, ObjPtr(key), delayMillis, IIf(IsNumeric(timerFunction), timerFunction, 0))
        If outTimerID = ObjPtr(key) Then
            Set this.timerIDs.item(key) = key 'store copy of key object so it doesn't go out of scope
            tryCreateTimer = True
        Else
            tryCreateTimer = False
        End If
    End If
    
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
        tryKillTimer = WinAPI.KillTimer(this.messageWindow.handle, timerID) 'kill the timer associated with that fn  & ID
        this.timerIDs.Remove timerID             'remove regardless of success, since it's either gone now, or was invalid anyway
    Else
        tryKillTimer = WinAPI.KillTimer(this.messageWindow.handle, timerID)
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

Private Sub log(ByVal loggerLevel As LogLevel, ByVal source As String, ByVal message As String)
    If Not LogManager.IsEnabled(loggerLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("TickerAPI" & loggerLevel, loggerLevel)
    End If
    LogManager.log loggerLevel, Toolbox.Strings.Format("{0} - {1}", source, message)
End Sub

Private Sub logError(ByVal source As String, ByVal errNum As Long, ByVal errDescription As String)
    If Not LogManager.IsEnabled(ErrorLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("TickerAPI", ErrorLevel)
    End If
    LogManager.log ErrorLevel, Toolbox.Strings.Format("{0} raised an error: #{1} - {2}", source, errNum, errDescription)
End Sub

Private Sub raiseError(ByVal errorCode As TimerError, Optional ByVal source As String = "raiseError")
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
    logError source, errorCode, Err.Description
    Err.Raise errorCode
End Sub

Private Sub checkIfLocked(Optional ByVal source As String = "checkIfLocked")
'    If this.isLocked Then
'        If this.timerIDs.Count = 0 Then          'message didn't come from some timer hanging around, so we can unlock safely
'            UnlockApi
'        Else                                     'there is a timer hanging around when this was called, so it *could* be responsible, so better safe than sorry
'            raiseError TimerApiLockedError, source
'        End If
'    End If
End Sub

Public Sub UnlockApi()
    this.isLocked = False
    log DebugLevel, "UnlockApi", "API unlocking called"
End Sub

Public Function tryScheduleUnlock() As Boolean
    'tryScheduleUnlock = tryScheduleProc(AddressOf TickerAPITimerProcs.UnlockCallbackProc)
End Function

Private Function setupMessageWindow() As Boolean
    'This should work whether or not a window & handle already exist
    If tryHookMessageHandler(AddressOf MessageWindowProcs.MessageWindowSubclassProc, this.messageWindow.name, this.messageWindow.handle) Then
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
    Set this.timerIDs = Cache.loadObject(TypeName(Me) & ".timerIDs", New Scripting.Dictionary)
    killAllTimers
    If setupMessageWindow Then
        'this should trigger unlock after the current synchronous call - whichever one triggered the predeclared class to initialise
        If this.isLocked Then tryScheduleUnlock
    Else
        this.isLocked = True
    End If
End Sub

Private Sub Class_Terminate()
    killAllTimers
End Sub

'For Testing
Friend Property Get messageWindowHandle() As LongPtr
    messageWindowHandle = this.messageWindow.handle
End Property

