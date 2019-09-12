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
    'TimerApiLockedError
    GetMessageWindowError
    InvalidTimerFunctionError
    [_End]
End Enum


'Private Const Default_Lock As Boolean = True
Private Const Default_Max_Timer_Count As Long = 100
'Private Const Default_Message_Window_Suffix As String = "MessageWindow"


Private Type tCallback
    
    'timerIDs As New Scripting.Dictionary         '{timerID: keyObj}
    'functionSet As New Dictionary
    '    recursionStack As New Stack
    '    isLocked As Boolean
    maxTimerCount As Long
    timerManager As ITimerManager
    timerSet As New TimerRepository
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
    raisePublicError GetMessageWindowError, loggerSourceName
    Resume                                       'for debugging - break above and jump to the error-raising statement
    
createTimerFail:
    logError "createTimer", Err.Number, Err.Description
    raisePublicError CreateTimerError, loggerSourceName
    Resume
    
scheduleProcFail:
    logError "scheduleProc", Err.Number, Err.Description
    KillTimerByID newTimerID                     'NOTE may raise an error
    raisePublicError SynchronousCallError, loggerSourceName
    Resume
    
End Function

Public Function StartManagedTimer(ByVal timerFunction As ITimerProc, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
    Const loggerSourceName As String = "StartManagedTimer"
    
    On Error GoTo getHandleFail
    Dim windowHandle As LongPtr
    windowHandle = getHandle
        
    On Error GoTo createTimerFail
    Dim newTimerID As LongPtr
    newTimerID = createTimer(windowHandle, ManagedCallbackWrapper.Create(timerFunction, data), delayMillis, AddressOf InternalTimerProcs.ManagedTimerCallbackInvoker)
    StartManagedTimer = newTimerID
    
    On Error GoTo scheduleProcFail
    If runImmediately Then
        'schedule immediate proc - No callback proc specified for managed timers
        If Not tryScheduleProc(windowHandle, newTimerID, AddressOf InternalTimerProcs.ManagedTimerCallbackInvoker) Then
            'queue is too full right now, no point scheduling as it wouldn't be evaluated in time anyway
            'could try flushing the queue instead
            log WarnLevel, loggerSourceName, "Message queue is too full to post to, so cannot runImmediately"
        End If
    End If
    
    log InfoLevel, loggerSourceName, printf("ManagedTimer with id {0} created", newTimerID)
    Exit Function
    
getHandleFail:
    logError "getHandle", Err.Number, Err.Description
    raisePublicError GetMessageWindowError, loggerSourceName
    Resume                                       'for debugging - break above and jump to the error-raising statement
    
createTimerFail:
    logError "createTimer", Err.Number, Err.Description
    raisePublicError CreateTimerError, loggerSourceName
    Resume
    
scheduleProcFail:
    logError "scheduleProc", Err.Number, Err.Description
    KillTimerByID newTimerID                     'NOTE may raise an error
    raisePublicError SynchronousCallError, loggerSourceName
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
        KillTimer windowHandle, timerID
            
        log InfoLevel, loggerSourceName, printf("Timer with id {0} destroyed", timerID)

    Else
        raisePublicError TimerNotFoundError, loggerSourceName
    End If
    
    Exit Sub
    
getHandleFail:
    logError "getHandle", Err.Number, Err.Description
    raisePublicError GetMessageWindowError, loggerSourceName
    Resume                                       'for debugging - break above and jump to the error-raising statement

killTimerFail:
    logError "killTimer", Err.Number, Err.Description
    raisePublicError DestroyTimerError, loggerSourceName
    Resume                                       'for debugging - break above and jump to the error-raising statement
    
End Sub

'@Description("Loops through all timers and kills those matching timerFunction - this can be a functionID, a functionObject(ITimerProc) or a functionName")
Public Sub KillTimersByFunction(ByVal timerFunction As Variant)
Attribute KillTimersByFunction.VB_Description = "Loops through all timers and kills those matching timerFunction - this can be a functionID, a functionObject(ITimerProc) or a functionName"
        
    On Error GoTo safeThrow
    If IsNumeric(timerFunction) Then
        If Int(timerFunction) = timerFunction Then 'not a decimal
            KillTimersByFunctionID timerFunction
        Else
            Err.Raise InvalidTimerFunctionError
        End If
    ElseIf IsObject(timerFunction) Then
        If TypeOf timerFunction Is ITimerProc Then
            KillTimersByFunctionID ObjPtr(timerFunction)
        Else
            Err.Raise InvalidTimerFunctionError
        End If
    ElseIf TypeName(timerFunction) = "String" Then
        KillTimersByFunctionName timerFunction
    Else
        Err.Raise InvalidTimerFunctionError
    End If
    
    Exit Sub
    
safeThrow:
    raisePublicError Err.Number, "KillTimersByFunction"
    Resume
    
End Sub

Public Sub KillTimersByFunctionID(ByVal timerFunctionID As LongPtr)

    On Error GoTo safeThrow
    Dim timerWrapper As ICallbackWrapper
    Dim timerID As Variant
    Dim killCount As Long
    For Each timerID In this.timerIDs.Keys
        Set timerWrapper = this.timerIDs.Item(timerID)
        If timerWrapper.FunctionID = timerFunctionID Then
            KillTimerByID timerID
            killCount = killCount + 1
        End If
    Next timerID
    
    Exit Sub
    
safeThrow:
    raisePublicError Err.Number, "KillTimersByFunctionID"
    Resume                                       'for debugging
End Sub

Public Sub KillTimersByFunctionName(ByVal timerFunctionName As String)
    On Error GoTo safeThrow
    Dim timerWrapper As ICallbackWrapper
    Dim timerID As Variant
    Dim killCount As Long
    For Each timerID In this.timerIDs.Keys
        Set timerWrapper = this.timerIDs.Item(timerID)
        If timerWrapper.FunctionName = timerFunctionName Then
            KillTimerByID timerID
            killCount = killCount + 1
        End If
    Next timerID
    
    Exit Sub
    
safeThrow:

    raisePublicError Err.Number, "KillTimersByFunctionName"
    Resume                                       'for debugging
End Sub

Public Sub raisePublicError(ByVal errorCode As TimerError, Optional ByVal Source As String = "raiseError")

    log TraceLevel, "raiseError", "Destroying timers and unsubclassing window so error can be raised"
    KillAllTimers                                'cleanly un subclasses message window as this prevents errors being raised
    
    Select Case errorCode
        Case TimerError.CreateTimerError
            Err.Description = "Couldn't create Timer"
            
        Case TimerError.DestroyTimerError
            Err.Description = "Uh Oh, can't kill the timer :("
            
        Case TimerError.GetMessageWindowError
            Err.Description = "Couldn't obtain a handle to the message window"
            
        Case TimerError.InvalidTimerFunctionError

            Err.Description = "Invalid timer function supplied; timer functions must be one of:" & vbNewLine _
                              & " - a TIMERPROC pointer" & vbNewLine _
                              & " - an ITimerProc instance" & vbNewLine _
                              & " - a class name String"
                  
        Case TimerError.SynchronousCallError
            Err.Description = "Error when running synchronously"
        
        Case TimerError.TimerNotFoundError
            Err.Description = "Timer not found"
            
        Case Else
            'rethrow error
            On Error Resume Next
            Err.Raise errorCode                  'fake raise to grab text for logging
            Dim errDescription As String
            errDescription = Err.Description
            On Error GoTo 0
            Err.Description = errDescription
            
    End Select
    
    logError Source, errorCode, Err.Description  'possibly overkill
    
    Err.Raise errorCode

    
End Sub

Private Sub terminateMessageWindow()
    'NOTE following an error, messageWindow is terminated but not reset
    If this.messageWindow Is Nothing Then
        log WarnLevel, "terminateMessageWindow", "Message window is not set"
    Else
        log DebugLevel, "terminateMessageWindow", "Destroying message window" '"Removing SUBCLASSPROCs: " & IIf(this.messageWindow.tryRemoveAllSubclasses, "successful", "unsuccessful")
        Set this.messageWindow = Nothing
    End If
End Sub

Private Sub Class_Initialize()
    'Set up defaults
    this.maxTimerCount = Default_Max_Timer_Count
    Set this.timerManager = New WindowsTimerManager
End Sub

Private Sub Class_Terminate()
    'To free up this resource we need to unsubclass the window. It will already be free if an error was just raised
    If Not this.messageWindow Is Nothing Then terminateMessageWindow
End Sub




