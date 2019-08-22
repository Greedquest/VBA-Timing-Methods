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

'Private Const Default_Lock As Boolean = True
Private Const Default_Max_Timer_Count As Long = 100
'Private Const Default_Message_Window_Suffix As String = "MessageWindow"
'
'Private Type tWindow
'    handle As LongPtr
'    name As String
'End Type

Private Type tCallback
    'BUG if you get rid of New here then subclassing the window screws synchronous error raising
    timerIDs As New Scripting.Dictionary             'timerID: keyObj
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
    
    Dim newTimerID As LongPtr
    If Not tryCreateTimer(UnmanagedCallbackWrapper.Create(timerFunction, data), delayMillis, newTimerID) Then
        raiseError CreateTimerError, loggerSourceName
    Else
        StartUnmanagedTimer = newTimerID
    End If
    If runImmediately Then
        'schedule immediate proc
        If Not tryScheduleProc(newTimerID, timerFunction) Then
            KillTimerByID newTimerID             'NOTE may raise error
        End If
    End If
End Function

Public Function StartManagedTimer(ByVal timerFunction As ITimerProc, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
    Const loggerSourceName As String = "StartManagedTimer"
    
    Dim newTimerID As LongPtr
    If Not tryCreateTimer(ManagedCallbackWrapper.Create(timerFunction, data), delayMillis, newTimerID) Then
        raiseError CreateTimerError, loggerSourceName
    Else
        StartManagedTimer = newTimerID
    End If
    If runImmediately Then
        'schedule immediate proc - No callback proc specified for managed timers
        If Not tryScheduleProc(newTimerID) Then
            KillTimerByID newTimerID             'NOTE may raise error
        End If
    End If
End Function

'@Description("API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimerByID(ByVal timerID As LongPtr)
Attribute KillTimerByID.VB_Description = "API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError"
    Const loggerSourceName As String = "KillTimerByID"
    
    If timerExists(timerID) Then
        If Not tryKillTimer(timerID) Then
            raiseError DestroyTimerError, loggerSourceName
        Else
            log InfoLevel, loggerSourceName, printf("Timer with id {0} destroyed", timerID)
        End If
    Else
        raiseError TimerNotFoundError, loggerSourceName
    End If
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
                If Not tryKillTimer(timerID) Then
                    raiseError DestroyTimerError, loggerSourceName
                    Exit For
                End If
                timerCount = timerCount + 1
            End If
        End If
    Next timerID
    log InfoLevel, loggerSourceName, printf("{0} timer{1} with function id {2} {3}", timerCount, IIf(timerCount = 1, vbNullString, "s"), timerFunction, IIf(timerCount = 0, "found", "destroyed"))
End Sub

'@Description("Sub loops through all known timers and tries to kill them. If none exist, it tries killing them by ID until none are left. Should not raise errors. Clients should just re-initialise API")
Friend Sub killAllTimers()
Attribute killAllTimers.VB_Description = "Sub loops through all known timers and tries to kill them. If none exist, it tries killing them by ID until none are left. Should not raise errors. Clients should just re-initialise API"
    Const loggerSourceName As String = "killAllTimers"
    'NOTE this procedure is called when raising errors so must not raise any itself
    On Error Resume Next
    log InfoLevel, loggerSourceName, printf("{0} registered timer(s)", this.timerIDs.Count)

    this.timerIDs.RemoveAll
    terminateMessageWindow
    setupMessageWindow

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
Private Function tryScheduleProc(ByVal timerID As LongPtr, Optional ByVal timerProc As LongPtr = 0) As Boolean
Attribute tryScheduleProc.VB_Description = "Posts a fake message to the message queue - timerProc should be 0 for managed timers"
    
    'Post fake message to queue
    tryScheduleProc = WinAPI.PostMessage(this.messageWindow.handle, WM_TIMER, timerID, timerProc) <> 0
        
End Function

'@Description("Creates a timer with specified delay and returns its ID. Takes a uniqie object which wraps data to pass to timers and also to use as a key")
Private Function tryCreateTimer(ByVal key As Object, ByVal delayMillis As Long, ByRef outTimerID As LongPtr) As Boolean
Attribute tryCreateTimer.VB_Description = "Creates a timer with specified delay and returns its ID. Takes an object to use as key, "
      
    On Error GoTo cleanFail
    
    'pass key as object since it's easy to downcast to ICallbackWrapper interface, but hard to up cast from ICallbackWrapper to every possible implementation
    Dim wrapper As ICallbackWrapper
    Set wrapper = key
    'key is either the TickerParams or the TimerProcWrapper so could upcast each
    
    'Create timer with ID of the key object; guaranteed to be unique
    'passing 0 as timerFunction will result in message posted to messageWindow - managed timer
    
    outTimerID = WinAPI.SetTimer(this.messageWindow.handle, ObjPtr(key), delayMillis, IIf(wrapper.CallbackType = cbRawPointer, wrapper.Callback, 0))
    If outTimerID = ObjPtr(key) Then 'confirm it has been made correctly
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
    tryKillTimer outTimerID 'NOTE clear up after error here as the outTimerID shouldn't be relied upon in caller to help clear up
    Resume cleanExit
    
End Function

Private Function tryKillTimer(ByVal timerID As LongPtr) As Boolean
    'TODO: could clear messages for this timer
    On Error GoTo cleanFail
    tryKillTimer = WinAPI.KillTimer(this.messageWindow.handle, timerID) <> 0 'Kill timer returns non-zero for success
    If timerExists(timerID) Then this.timerIDs.Remove timerID             'remove regardless of success, since it's either gone now, or was invalid anyway

cleanExit:
    Exit Function
    
cleanFail:
    logError "tryKillTimer", Err.Number, Err.Description
    tryKillTimer = False
    Resume cleanExit
    
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
'        Case TimerError.TimerApiLockedError
'            Err.Description = "Call the Unlock method before attempting to use the API"
        Case Else
            'rethrow error
    End Select
    logError Source, errorCode, Err.Description
    Err.Raise errorCode
End Sub

Private Sub setupMessageWindow()
    If Not ModelessMessageWindow.tryCreate(this.messageWindow, AddressOf MessageWindowProcs.ManagedTimerMessageWindowSubclassProc) Then
        log FatalLevel, "setupMessageWindow", "Unable to hook message window"
    End If
End Sub

Private Sub terminateMessageWindow()
    Set this.messageWindow = Nothing
End Sub

Private Sub Class_Initialize()
    'Set up defaults
    this.maxTimerCount = Default_Max_Timer_Count
    setupMessageWindow
End Sub

Private Sub Class_Terminate()
    terminateMessageWindow
End Sub

'For Testing
Friend Property Get messageWindowHandle() As LongPtr
    messageWindowHandle = this.messageWindow.handle
End Property

Friend Sub makeBadness()
    Debug.Print "Un sub-classing - ";
    Debug.Print this.messageWindow.tryRemoveAllSubclasses
    Debug.Print "Raising Error"
    Err.Raise 5
    Debug.Print "Error Raised"
End Sub
