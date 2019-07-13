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
'@Folder("API")
'@PredeclaredID
'@ModuleDescription("API for setting up timers to callback functions, wraps Windows Timers")
Option Explicit

''@Description("Windows Timer Message https://docs.microsoft.com/windows/desktop/winmsg/wm-timer")
Public Enum WindowsMessage
    WM_TIMER = &H113
    WM_NOTIFY = &H4E                             'arbitrary, sounds nice though
End Enum

Public Enum TimerError
    CreateTimerError = vbObjectError + 1
    DestroyTimerError
    TimerNotFoundError
    SynchronousCallError
    TimerApiLockedError
End Enum

Private Const Default_Lock As Boolean = True

Private Declare Function ApiSetTimer Lib "user32" Alias "SetTimer" ( _
                         ByVal HWnd As Long, _
                         ByVal nIDEvent As Long, _
                         ByVal uElapse As Long, _
                         ByVal lpTimerFunc As Long) As Long

Private Declare Function ApiKillTimer Lib "user32" Alias "KillTimer" ( _
                         ByVal HWnd As Long, _
                         ByVal nIDEvent As Long) As Long

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                         ByVal lpPrevWndFunc As Long, _
                         ByRef timerFlag As Bool, _
                         Optional ByVal message As WindowsMessage = WM_NOTIFY, _
                         Optional ByVal timerID As Long = 0, _
                         Optional ByVal unused3 As Long) As Long

Private Type tCallback
    timerIDs As New Dictionary
    functionSet As New Dictionary
    recursionStack As New Stack
    isLocked As Boolean
End Type

Private this As tCallback

'@Description("Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error")
Public Function StartTimer(ByVal timerFunction As Long, Optional ByVal runImmediately As Boolean = True, Optional ByVal delayMillis As Long = 500) As Long
Attribute StartTimer.VB_Description = "Create new timer instance with optional synchronous first call. Returns the ID of the newly created windows timer. Can raise SynchronousCallError if timerFunction fails (and is trapped - unlikely). Raises CreateTimerError if there is an API error"
    'If outTimerID = -1 Then outTimerID = getUniqueID(timerFunction)
    'Set this.recursionStack = New Stack
    If runImmediately Then
        Dim makeTimer As Bool
        Set makeTimer = Bool.Create(True)        'TODO can raise error
        this.recursionStack.Push makeTimer
        If Not tryCallProc(timerFunction, makeTimer) Then 'NOTE this will always be true as errors are untrappable and pointer is assumed valid
            raiseError SynchronousCallError, "StartTimer"
        End If
        'Check whether timerFunction has edited the flag on the top of the stack - could do a 1 liner
        Set makeTimer = this.recursionStack.Pop
        If Not makeTimer Then Exit Function
    End If
    Dim newTimerID As Long
    If Not tryCreateTimer(timerFunction, newTimerID, delayMillis) Then
        raiseError CreateTimerError, "StartTimer"
    Else
        StartTimer = newTimerID
    End If
End Function

'@Description("API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimerByID(ByVal timerID As Long)
Attribute KillTimerByID.VB_Description = "API kills windows timer on this handle by ID. Unregistered ID raises TimerNotFoundError, failure to destroy a registered ID raises DestroyTimerError"
    If timerExists(timerID) Then
        If Not tryKillTimer(timerID) Then
            raiseError DestroyTimerError, "KillTimerByID"
        End If
    Else
        raiseError TimerNotFoundError, "KillTimerByID"
    End If
End Sub

'@Description("API kills windows timers registered to a certain callback. No error is raised if none are found,failure to destroy a registered ID raises DestroyTimerError")
Public Sub KillTimersByFunction(ByVal timerFunction As Long)
Attribute KillTimersByFunction.VB_Description = "API kills windows timers registered to a certain callback. No error is raised if none are found,failure to destroy a registered ID raises DestroyTimerError"
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

'@Description("Sub loops through all known timers and tries to kill them. If none exist, it tries killing them by ID until none are left. Should not raise errors")
Public Sub KillAllTimers()
Attribute KillAllTimers.VB_Description = "Sub loops through all known timers and tries to kill them. If none exist, it tries killing them by ID until none are left. Should not raise errors"
    'NOTE this procedure is called when raising errors so must not raise any itself
    log DebugLevel, "KillAllTimers", printf("{0} registered timer(s)", this.timerIDs.Count)
    If this.timerIDs.Count = 0 Then
        Dim i As Long
        log DebugLevel, "KillAllTimers", "Looping through and killing possible timers on window handle"
        Do                                       'loop through all timers for this hwnd
            i = i + 1
        Loop While tryKillTimer(i) Or i < 100
    Else
        Dim timerID As Variant
        For Each timerID In this.timerIDs
            On Error Resume Next                 'no point returning success since there's nothing we can do about it
            ApiKillTimer 0, this.timerIDs(timerID)
            tryKillTimer this.timerIDs(timerID)
            log DebugLevel, "KillAllTimers", printf("Killing {0}", timerID)
        Next timerID
        this.timerIDs.RemoveAll
    End If
End Sub

'@Description("Checks the runningSet for the item with key timerFunction, returns the item if found, or false otherwise")
Private Property Get timerExists(ByVal timerID As Long) As Boolean
Attribute timerExists.VB_Description = "Checks the runningSet for the item with key timerFunction, returns the item if found, or false otherwise"
    timerExists = this.timerIDs.Exists(timerID)
End Property

Private Function tryCallProc(ByVal timerFunction As Long, ByRef timerFlag As Bool) As Boolean
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

Private Function tryCreateTimer(ByVal timerFunction As Long, ByRef outTimerID As Long, ByVal delayMillis As Long) As Boolean
    On Error GoTo cleanFail
    If Not timerExists(outTimerID) Then          'only make a new timer if there isn't one already
        outTimerID = ApiSetTimer(Application.HWnd, getUniqueID, delayMillis, timerFunction)
        this.timerIDs(outTimerID) = timerFunction
        tryCreateTimer = True
    End If
    
cleanExit:
    Exit Function
    
cleanFail:
    logError "tryCreateTimer", Err.Number, Err.Description
    tryCreateTimer = False
    Resume cleanExit
    
End Function

Private Function tryKillTimer(ByVal timerID As Long) As Boolean
    On Error GoTo cleanFail
    If timerExists(timerID) Then
        tryKillTimer = ApiKillTimer(Application.HWnd, timerID) 'kill the timer associated with that fn  & ID
        this.timerIDs.Remove timerID             'remove regardless of success, since it's either gone now, or was invalid anyway
    Else
        tryKillTimer = ApiKillTimer(Application.HWnd, timerID)
    End If

cleanExit:
    Exit Function
    
cleanFail:
    logError "tryKillTimer", Err.Number, Err.Description
    tryKillTimer = False
    Resume cleanExit
    
End Function

Private Function getUniqueID() As Long
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
    KillAllTimers
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
    If this.isLocked Then raiseError TimerApiLockedError, source
End Sub

Public Sub UnlockTimer()
    this.isLocked = False
End Sub

Private Sub Class_Initialize()
    this.isLocked = Default_Lock
    KillAllTimers
End Sub

Private Sub Class_Terminate()
    KillAllTimers
End Sub


