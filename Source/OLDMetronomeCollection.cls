VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OLDMetronomeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@Folder("OLDSecondLevelAPI")
'@IgnoreModule

Option Explicit
'@PredeclaredId

Private Type TTimer
    activeTickers As New Dictionary
    synchronousTickers As New Stack
End Type

Private this As TTimer

Public Function Add(ByVal tickInterval As Long, ByVal runFirstTickSynchronously As Boolean, Optional ByVal autoExpireTickCount As Long = InfiniteTicks) As OLDMetronome

    With New OLDMetronome
        .Controls.tickIntervalMillis = tickInterval
        .Controls.remainingTicks = autoExpireTickCount
        .Controls.isFirstTickSynchronous = runFirstTickSynchronously
        Set Add = .Controls                      'UpCast IMetronomeControls to Metronome
    End With
    
End Function

Friend Sub CallbackNotify(ByVal createTimer As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
    Dim ticker As OLDIMetronomeControls
    log DebugLevel, "CallbackNotify", printf("Callback called with ID {0}", timerID)
    If Not tryGetTickerFromID(timerID, ticker) Then
        'Rogue timerID shouldn't be messaging our private callback! Disable it
        If message = WM_NOTIFY Then
            Err.Raise 5, Description:="Unreacheable"
            'Synchronous call, dereference flag and set it to false so no timer is made
            'TickerParams.FromPtr(createTimer).TickerIsRunning = False
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

Friend Function tryInitialiseTimer(ByVal ticker As OLDIMetronomeControls) As Boolean
    
    'synchronous ticks
    If ticker.isFirstTickSynchronous Then
        this.synchronousTickers.Push ticker
    End If
    
    On Error GoTo CleanFail
    ticker.timerID = TickerAPI.StartUnmanagedTimer(AddressOf WaitableTimerCallbackProc, ticker.isFirstTickSynchronous, ticker.tickIntervalMillis)
    this.activeTickers.Add key:=ticker.timerID, Item:=ticker
    tryInitialiseTimer = True
        
CleanExit:
    Exit Function
    
CleanFail:
    logError "tryInitialiseTimer", Err.Number, Err.Description
    tryInitialiseTimer = False
    Resume CleanExit
    
End Function

Friend Function tryTerminateTimer(ByVal ticker As OLDIMetronomeControls) As Boolean
    On Error GoTo CleanFail
    TickerAPI.KillTimerByID ticker.timerID
    this.activeTickers.Remove ticker.timerID
    tryTerminateTimer = True
    
CleanExit:
    Exit Function
    
CleanFail:
    logError "tryTerminateTimer", Err.Number, Err.Description
    tryTerminateTimer = False
    Resume CleanExit
    
End Function

Private Function tryGetTickerFromID(ByVal timerID As Long, ByRef outTicker As OLDIMetronomeControls) As Boolean
    
    Const synchronousTimerID As Long = 0         'NOTE this should really be defined in TickerAPI TryCallProc or another public place
    
    'it's possible we are in a synchronous stack - the ticker for a given timer won't have an ID corresponding to a timer yet
    If timerID = synchronousTimerID Then
        If this.synchronousTickers.Count = 0 Then
            tryGetTickerFromID = False
        Else
            Set outTicker = this.synchronousTickers.Pop
            tryGetTickerFromID = True
        End If
        
    ElseIf Not this.activeTickers.Exists(timerID) Then
        tryGetTickerFromID = False
        
    Else
        Set outTicker = this.activeTickers.Item(timerID)
        tryGetTickerFromID = True
    End If
End Function

Private Sub Class_Terminate()
    TickerAPI.KillTimersByFunction AddressOf WaitableTimerCallbackProc
End Sub
