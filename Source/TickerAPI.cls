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
'@PredeclaredID: To ensure it's a singleton in other projects and avoid async nulling
'@ModuleDescription("API for setting up timers to callback functions, wraps Windows Timers")
Option Explicit

Public Enum TimerError
    [_ErrBase] = 0
    [_Start] = vbObjectError + [_ErrBase]
    CreateTimerError
    DestroyTimerError
    TimerNotFoundError
    SynchronousCallError
    InvalidTimerFunctionError
    GenerateTimerDataError
    [_End]
End Enum

Private Const Default_Max_Timer_Count As Long = 100

Private Type tCallback
    maxTimerCount As Long
    timerManager As ITimerManager
    timerDataRepo As New TimerRepository
End Type

Private this As tCallback

Private Sub Class_Initialize()
    'Set up defaults
    this.maxTimerCount = Default_Max_Timer_Count
    Set this.timerManager = New WindowsTimerManager
End Sub

'@Description("Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error")
Public Function StartUnmanagedTimer(ByVal timerFunction As LongPtr, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
Attribute StartUnmanagedTimer.VB_Description = "Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error"
    Const loggerSourceName As String = "StartUnmanagedTimer"
   
    'On Error GoTo generateTimerDataFail
    Dim timerInfo As TimerData
    If Not TryGenerateTimerData(timerFunction, data, timerInfo) Then raisePublicError GenerateTimerDataError, loggerSourceName
        
    On Error GoTo createTimerFail
    this.timerManager.StartTimer timerInfo, delayMillis
    StartUnmanagedTimer = timerInfo.ID
    
    On Error GoTo scheduleProcFail
    If runImmediately Then
        If Not this.timerManager.TryTriggerTimer(timerInfo) Then
            'queue is too full right now, no point scheduling as it wouldn't be evaluated in time anyway
            'could try flushing the queue instead
            log WarnLevel, loggerSourceName, "Message queue is too full to post to, so cannot runImmediately"
        End If
    End If
    
    log InfoLevel, loggerSourceName, printf("UnmanagedTimer with id {0} created", timerInfo.ID)
    Exit Function
    
generateTimerDataFail:
    logError "timerSet.Add", Err.Number, Err.Description
    raisePublicError GenerateTimerDataError, loggerSourceName
    Resume                                       'for debugging - break above and jump to the error-raising statement
    
createTimerFail:
    logError "createTimer", Err.Number, Err.Description
    this.timerDataRepo.Remove timerInfo
    raisePublicError CreateTimerError, loggerSourceName
    Resume
    
scheduleProcFail:
    logError "scheduleProc", Err.Number, Err.Description
    KillTimerByID timerInfo.ID                   'NOTE may raise its own public error
    raisePublicError SynchronousCallError, loggerSourceName
    Resume
    
End Function
Private Function TryGenerateTimerData(ByVal timerFunction As Variant, ByVal data As Variant, ByRef outTimerInfo As TimerData) As Boolean
    Const loggerSourceName As String = "TryGenerateTimerData"
    
    On Error GoTo CleanFail
    If IsObject(timerFunction) Then
        Set outTimerInfo = TimerData.Create(ManagedCallbackWrapper.Create(timerFunction, data))
    ElseIf VarType(timerFunction) = vbLongPtr Then
        Set outTimerInfo = TimerData.Create(UnmanagedCallbackWrapper.Create(timerFunction, data))
    Else
        Err.Description = "Unreachable: invalid timerFunctionType"
        Err.Raise 13
    End If
    this.timerDataRepo.Add outTimerInfo
    TryGenerateTimerData = True
    
CleanExit:
    Exit Function
    
CleanFail:
    logError loggerSourceName, Err.Number, Err.Description
    TryGenerateTimerData = False
    Resume CleanExit ' DEBUG: comment-out this statement
    Stop
    Resume
End Function

Public Function StartManagedTimer(ByVal timerFunction As ITimerProc, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
    Const loggerSourceName As String = "StartManagedTimer"
    
    Dim timerInfo As TimerData
    If Not TryGenerateTimerData(timerFunction, data, timerInfo) Then raisePublicError GenerateTimerDataError, loggerSourceName
     
    On Error GoTo createTimerFail
    this.timerManager.StartTimer timerInfo, delayMillis
    StartManagedTimer = timerInfo.ID
    
    On Error GoTo scheduleProcFail
    If runImmediately Then
        If Not this.timerManager.TryTriggerTimer(timerInfo) Then
            'queue is too full right now, no point scheduling as it wouldn't be evaluated in time anyway
            'could try flushing the queue instead
            log WarnLevel, loggerSourceName, "Message queue is too full to post to, so cannot runImmediately"
        End If
    End If
    
    log InfoLevel, loggerSourceName, printf("ManagedTimer with id {0} created", timerInfo.ID)
    Exit Function
    
generateTimerDataFail:
    logError "timerSet.Add", Err.Number, Err.Description
    raisePublicError GenerateTimerDataError, loggerSourceName
    Resume                                       'for debugging - break above and jump to the error-raising statement
    
createTimerFail:
    logError "createTimer", Err.Number, Err.Description
    this.timerDataRepo.Remove timerInfo
    raisePublicError CreateTimerError, loggerSourceName
    Resume
    
scheduleProcFail:
    logError "scheduleProc", Err.Number, Err.Description
    KillTimerByID timerInfo.ID                   'NOTE may raise an error
    raisePublicError SynchronousCallError, loggerSourceName
    Resume
    
End Function

'@Description("API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimerByID(ByVal timerID As LongPtr)
Attribute KillTimerByID.VB_Description = "API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError"
    Const loggerSourceName As String = "KillTimerByID"
    
    If this.timerDataRepo.Exists(timerID) Then
          
        On Error GoTo killTimerFail
        Dim timerInfo As TimerData
        Set timerInfo = this.timerDataRepo.Item(timerID)
        
        this.timerDataRepo.Remove timerInfo
        this.timerManager.KillTimer timerInfo
          
        log InfoLevel, loggerSourceName, printf("Timer with id {0} destroyed", timerInfo.ID)

    Else
        raisePublicError TimerNotFoundError, loggerSourceName
    End If
    
    Exit Sub
    
killTimerFail:
    logError "killTimer", Err.Number, Err.Description
    raisePublicError DestroyTimerError, loggerSourceName
    Resume                                       'for debugging - break above and jump to the error-raising statement
    
End Sub

'@Description("Loops through all timers and kills those matching timerFunction - this can be a functionID, a functionObject(ITimerProc) or a functionName")
Public Sub KillTimersByFunction(ByVal timerFunction As Variant)
Attribute KillTimersByFunction.VB_Description = "Loops through all timers and kills those matching timerFunction - this can be a functionID, a functionObject(ITimerProc) or a functionName"
    Const errorSourceName As String = "KillTimersByFunction"
    
    'REVIEW slightly nasty how this function calls public methods which may raise public errors :/
    'Could just remove error guard and raise them itself, but that's risky as there might be unhandled internal errors
    On Error GoTo safeThrow
    If IsNumeric(timerFunction) Then
        If Int(timerFunction) = timerFunction Then 'not a decimal
            Me.KillTimersByFunctionID timerFunction
        Else
            raisePublicError InvalidTimerFunctionError, errorSourceName
        End If
        
    ElseIf IsObject(timerFunction) Then
        If TypeOf timerFunction Is ITimerProc Then
            Me.KillTimersByFunctionID ObjPtr(timerFunction)
        Else
            raisePublicError InvalidTimerFunctionError, errorSourceName
        End If
        
    ElseIf TypeName(timerFunction) = "String" Then
        Me.KillTimersByFunctionName timerFunction
    Else
        raisePublicError InvalidTimerFunctionError, errorSourceName
    End If
    
    Exit Sub
    
safeThrow:
    
    'check if within custom error range; if so then don't rethrow as that would re-terminate and double log the error
    If Err.Number > TimerError.[_End] Or Err.Number < TimerError.[_Start] Then
        'Unexpected Errors: must throw them to public; no sense condensing as these are all unexpected
        raisePublicError Err.Number, "KillTimersByFunction"
    Else
        'Public Errors: all the cleanup is done, safe to just re-throw
        Err.Raise Err.Number
    End If
    Resume
    
End Sub

Public Sub KillTimersByFunctionID(ByVal timerFunctionID As LongPtr)

    On Error GoTo safeThrow

    Dim timer As TimerData
    For Each timer In this.timerDataRepo.FilterByFunctionID(timerFunctionID)
        KillTimerByID timer.ID
    Next timer
    
    Exit Sub
    
safeThrow:
    raisePublicError Err.Number, "KillTimersByFunctionID"
    Resume                                       'for debugging
    
End Sub

Public Sub KillTimersByFunctionName(ByVal timerFunctionName As String)
    On Error GoTo safeThrow

    Dim timer As TimerData
    For Each timer In this.timerDataRepo.FilterByFunctionName(timerFunctionName)
        KillTimerByID timer.ID
    Next timer
    
    Exit Sub
    
safeThrow:
    raisePublicError Err.Number, "KillTimersByFunctionName"
    Resume                                       'for debugging
End Sub

Public Sub KillAll()
    'NOTE this is called when raising errors so must not generate any itself
    On Error Resume Next
    this.timerManager.KillAllTimers this.timerDataRepo.ToArray
    this.timerDataRepo.Clear
    If Err.Number <> 0 Then logError "KillAll", Err.Number, Err.Description
    On Error GoTo 0
End Sub

Private Sub raisePublicError(ByVal errorCode As TimerError, Optional ByVal Source As String = "raiseError")

    log TraceLevel, "raiseError", "Destroying timers so error can be raised"
    Me.KillAll
    
    Select Case errorCode
        Case TimerError.CreateTimerError
            Err.Description = "Couldn't create Timer"
            
        Case TimerError.DestroyTimerError
            Err.Description = "Uh Oh, can't kill the timer :("
           
        Case TimerError.GenerateTimerDataError
            Err.Description = "Unable to add/retrieve timer data from the repository"
            
        Case TimerError.InvalidTimerFunctionError

            Err.Description = "Invalid timer function supplied; timer functions must be one of:" & vbNewLine _
                              & " - a TIMERPROC or ITimerProc pointer" & vbNewLine _
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
    
    logError Source, errorCode, Err.Description  'REVIEW possibly overkill
    
    Err.Raise errorCode

End Sub

'For testing
Friend Property Get messageWindowHandle() As LongPtr
    'only on windows
    Dim timerManager As WindowsTimerManager
    Set timerManager = this.timerManager
    messageWindowHandle = timerManager.messageWindowHandle
End Property
