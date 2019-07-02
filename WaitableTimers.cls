VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WaitableTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("TimerAPI")

Option Explicit
'@PredeclaredId

Private Type TTimer
    timers As New Collection
End Type

Private this As TTimer


Public Function Add(ByVal tickInterval As Long, ByVal startImmediately As Boolean, ByVal autoExpireTickCount As Long) As Metronome

End Function

Public Sub CallbackNotify(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Dim ticker As IMetronomeControls
    If Not tryGetTickerFromID(timerID, ticker) Then
        TickerAPI.KillTimerByID timerID
    End If
    
    If message = WM_NOTIFY Then
        Bool.FromPtr(createTimer) = False
    Else
        TickerAPI.KillTimerByID timerID
    End If
End Sub

Public Function tryStartTimer(ticker As IMetronomeControls) As Boolean
    On Error GoTo cleanFail
    ticker.isRunning = True
    ticker.timerID = TickerAPI.StartTimer(AddressOf WaitableTimerCallbackProc, ticker.isFirstTickSynchronous, ticker.tickIntervalMillis)
    tryStartTimer = True
        
cleanExit:
    Exit Function
    
cleanFail:
    logError "tryStartTimer", Err.Number, Err.Description
    tryStartTimer = False
    Resume cleanExit
    
End Function

Public Function tryPauseTimer(ticker As IMetronomeControls) As Boolean
    On Error GoTo cleanFail
    TickerAPI.KillTimerByID ticker.timerID
    ticker.isRunning = False
    tryPauseTimer = True
    
cleanExit:
    Exit Function
    
cleanFail:
    logError "tryPauseTimer", Err.Number, Err.Description
    tryPauseTimer = False
    Resume cleanExit
    
End Function

Public Function tryGetTickerFromID(ByVal timerID As Long, ByRef outTicker As IMetronomeControls) As Boolean
    Const NO_ITEM_ERROR As Long = 5
    On Error Resume Next
    Set outTicker = this.timers.item(timerID)
    Select Case Err.Number
        Case 0 'No Errors
            tryGetTickerFromID = True
            
        Case NO_ITEM_ERROR 'No ticker for that ID
            tryGetTickerFromID = False
            
        Case Else 'Unexpected error
            tryGetTickerFromID = False
            logError "tryGetTickerFromID", Err.Number, Err.Description
            
    End Select
    
End Function

Private Sub logError(ByVal source As String, ByVal errNum As Long, ByVal errDescription As String)
    If Not logmanager.IsEnabled(ErrorLevel) Then 'check a logger is registered
        logmanager.Register DebugLogger.Create(TypeName(Me), ErrorLevel)
    End If
    logmanager.log ErrorLevel, Toolbox.Strings.Format("{0} raised an error: #{1} - {2}", source, errNum, errDescription)
End Sub

Private Sub log(ByVal loggerLevel As LogLevel, ByVal source As String, ByVal message As String)
    If Not logmanager.IsEnabled(loggerLevel) Then 'check a logger is registered
        logmanager.Register DebugLogger.Create(TypeName(Me) & loggerLevel, loggerLevel)
    End If
    logmanager.log loggerLevel, Toolbox.Strings.Format("{0} - {1}", source, message)
End Sub

Private Sub Class_Terminate()
    TickerAPI.KillTimersByFunction AddressOf WaitableTimerCallbackProc
End Sub
