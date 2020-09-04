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

Public Function StartUnmanagedTimer(ByVal timerFunction As LongPtr, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
    Const loggerSourceName As String = "StartUnmanagedTimer"
    StartUnmanagedTimer = StartTimer(loggerSourceName, timerFunction, runImmediately, delayMillis, data)
End Function

'@Ignore FunctionReturnValueAlwaysDiscarded:API
Public Function StartManagedTimer(ByVal timerFunction As ITimerProc, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500, Optional ByVal data As Variant) As LongPtr
    Const loggerSourceName As String = "StartManagedTimer"
    StartManagedTimer = StartTimer(loggerSourceName, timerFunction, runImmediately, delayMillis, data)
End Function

'@Description("Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error")
Private Function StartTimer(ByVal loggerSourceName As String, ByVal timerFunction As Variant, ByVal runImmediately As Boolean, ByVal delayMillis As Long, ByVal data As Variant) As LongPtr
Attribute StartTimer.VB_Description = "Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error"
    
    Dim timerInfo As TimerData
    If Not TryGenerateTimerData(timerFunction, data, timerInfo) Then raisePublicError GenerateTimerDataError, loggerSourceName
    If Not TryCreateTimer(timerInfo, delayMillis, StartTimer) Then raisePublicError CreateTimerError, loggerSourceName
    If runImmediately Then
        If Not AttemptToTriggerTimer(timerInfo) Then raisePublicError SynchronousCallError, loggerSourceName
    End If
    
    logMessage InfoLevel, loggerSourceName, printf("Timer with id {0} created", timerInfo.ID)
End Function

Private Function TryGenerateTimerData(ByVal timerFunction As Variant, ByVal data As Variant, ByRef outTimerInfo As TimerData) As Boolean
    Const loggerSourceName As String = "TryGenerateTimerData"
    
    On Error GoTo CleanFail
    If IsObject(timerFunction) Then
        Set outTimerInfo = TimerData.Create(ManagedCallbackWrapper.Create(timerFunction, data))
    ElseIf VarType(timerFunction) = vbLongPtr Then
        Set outTimerInfo = TimerData.Create(UnmanagedCallbackWrapper.Create(timerFunction, data))
    Else                                         'NOTE: Unreachable since timerFunction is strongly typed in the parent procedures
        Err.Description = "Invalid timerFunctionType"
        Err.Raise 13
    End If
    this.timerDataRepo.Add outTimerInfo
    TryGenerateTimerData = True
    
CleanExit:
    Exit Function
    
CleanFail:
    logError loggerSourceName, Err.Number, Err.Description
    TryGenerateTimerData = False
    Resume CleanExit                             'DEBUG: comment-out this statement
Stop
    Resume
End Function

Private Function TryCreateTimer(ByVal timerInfo As TimerData, ByVal delayMillis As Long, ByRef outTimerID As LongPtr) As Boolean
    Const loggerSourceName As String = "TryCreateTimer"
    
    On Error GoTo CleanFail
    this.timerManager.StartTimer timerInfo, delayMillis
    outTimerID = timerInfo.ID
    TryCreateTimer = True
    
CleanExit:
    Exit Function
    
CleanFail:
    logError loggerSourceName, Err.Number, Err.Description
    this.timerDataRepo.Remove timerInfo
    TryCreateTimer = False
    Resume CleanExit                             'DEBUG: comment-out this statement
Stop
    Resume
End Function

Private Function AttemptToTriggerTimer(ByVal timerInfo As TimerData) As Boolean
    Const loggerSourceName As String = "AttemptToTriggerTimer"
    
    On Error GoTo CleanFail
    If Not this.timerManager.TryTriggerTimer(timerInfo) Then
        logMessage WarnLevel, loggerSourceName, "Attempt failed with no error"
    End If
    AttemptToTriggerTimer = True
    
CleanExit:
    Exit Function
    
CleanFail:
    KillTimerByID timerInfo.ID
    logError loggerSourceName, Err.Number, Err.Description
    AttemptToTriggerTimer = False
    Resume CleanExit                             'DEBUG: comment-out this statement
Stop
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
          
        logMessage InfoLevel, loggerSourceName, printf("Timer with id {0} destroyed", timerInfo.ID)

    Else
        raisePublicError TimerNotFoundError, loggerSourceName
    End If

CleanExit:
    Exit Sub
    
killTimerFail:
    logError "killTimer", Err.Number, Err.Description
    raisePublicError DestroyTimerError, loggerSourceName
    Resume CleanExit                             'for debugging - break above and jump to the error-raising statement
Stop
    Resume
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

    Dim matchingTimer As TimerData
    For Each matchingTimer In this.timerDataRepo.FilterByFunctionID(timerFunctionID)
        KillTimerByID matchingTimer.ID
    Next matchingTimer
    
    Exit Sub
    
safeThrow:
    raisePublicError Err.Number, "KillTimersByFunctionID"
    Resume                                       'for debugging
    
End Sub

Public Sub KillTimersByFunctionName(ByVal timerFunctionName As String)
    On Error GoTo safeThrow

    Dim matchingTimer As TimerData
    For Each matchingTimer In this.timerDataRepo.FilterByFunctionName(timerFunctionName)
        KillTimerByID matchingTimer.ID
    Next matchingTimer
    
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

    logMessage TraceLevel, "raiseError", "Destroying timers so error can be raised safely"
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
