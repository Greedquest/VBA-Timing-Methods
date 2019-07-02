VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetronomeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("TimerAPI")

Option Explicit
'@PredeclaredId

Private Type TTimer
    activeTickers As New Dictionary
End Type

Private this As TTimer

Public Function Add(ByVal tickInterval As Long, ByVal runFirstTickSynchronously As Boolean, Optional ByVal autoExpireTickCount As Long = InfiniteTicks) As Metronome

    With New Metronome
        .Controls.tickIntervalMillis = tickInterval
        .Controls.remainingTicks = autoExpireTickCount
        .Controls.isFirstTickSynchronous = runFirstTickSynchronously
        Set Add = .Controls                      'UpCast IMetronomeControls to Metronome
    End With
    
End Function

Friend Sub CallbackNotify(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Dim ticker As IMetronomeControls
    If Not tryGetTickerFromID(timerID, ticker) Then
        'Rogue timerID shouldn't be messaging our private callback! Disable it
        If message = WM_NOTIFY Then
            'Synchronous call, dereference flag and set it to false so no timer is made
            Bool.FromPtr(createTimer) = False
        Else
            'Timer call, destroy the timer
            On Error Resume Next
            TickerAPI.KillTimerByID timerID
            If Err.Number <> 0 Then logError "CallbackNotify", Err.Number, Err.Description
            On Error GoTo 0
        End If
        Exit Sub
    End If
    'Good callback, do we want to trigger a tick event?
    If ticker.remainingTicks = InfiniteTicks Then
        ticker.Tick
    Else
        ticker.remainingTicks = ticker.remainingTicks - 1
        If ticker.remainingTicks = 0 Then
            ticker.isRunning = False
        End If
        ticker.Tick
    End If
    
End Sub

Public Function tryInitialiseTimer(ticker As IMetronomeControls) As Boolean
    On Error GoTo cleanFail
    ticker.timerID = TickerAPI.StartTimer(AddressOf WaitableTimerCallbackProc, ticker.isFirstTickSynchronous, ticker.tickIntervalMillis)
    this.activeTickers.Add key:=ticker.timerID, item:=ticker
    tryInitialiseTimer = True
        
cleanExit:
    Exit Function
    
cleanFail:
    logError "tryInitialiseTimer", Err.Number, Err.Description
    tryInitialiseTimer = False
    Resume cleanExit
    
End Function

Public Function tryTerminateTimer(ticker As IMetronomeControls) As Boolean
    On Error GoTo cleanFail
    TickerAPI.KillTimerByID ticker.timerID
    this.activeTickers.Remove ticker.timerID
    tryTerminateTimer = True
    
cleanExit:
    Exit Function
    
cleanFail:
    logError "tryTerminateTimer", Err.Number, Err.Description
    tryTerminateTimer = False
    Resume cleanExit
    
End Function

Public Function tryGetTickerFromID(ByVal timerID As Long, ByRef outTicker As IMetronomeControls) As Boolean

    If Not this.activeTickers.Exists(timerID) Then
        tryGetTickerFromID = False
    Else
        Set outTicker = this.activeTickers.item(timerID)
        tryGetTickerFromID = True
    End If
    
End Function

Private Sub logError(ByVal source As String, ByVal errNum As Long, ByVal errDescription As String)
    If Not LogManager.IsEnabled(ErrorLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create(TypeName(Me), ErrorLevel)
    End If
    LogManager.log ErrorLevel, Toolbox.Strings.Format("{0} raised an error: #{1} - {2}", source, errNum, errDescription)
End Sub

Private Sub log(ByVal loggerLevel As LogLevel, ByVal source As String, ByVal message As String)
    If Not LogManager.IsEnabled(loggerLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create(TypeName(Me) & loggerLevel, loggerLevel)
    End If
    LogManager.log loggerLevel, Toolbox.Strings.Format("{0} - {1}", source, message)
End Sub

Private Sub Class_Terminate()
    TickerAPI.KillTimersByFunction AddressOf WaitableTimerCallbackProc
End Sub

