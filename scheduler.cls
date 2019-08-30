VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "scheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests")
Option Explicit

Private WithEvents log As testLog
Attribute log.VB_VarHelpID = -1
Public Enum loggerEventType
    loggerErrorEvent = 2 ^ 0
    loggerLogEvent = 2 ^ 1
    loggerAllEvents = 2 ^ 2
End Enum

Private errorHappened As Boolean
Private logHappened As Boolean

Private Sub log_afterError()
    errorHappened = True
End Sub

Private Sub log_afterLog()
    logHappened = True
End Sub

Public Sub doEventsWait(ByVal logger As testLog, Optional ByVal delayMillis As Long = defaultTimerDelay, Optional ByVal trigger As loggerEventType = loggerAllEvents)
    Set logger = log
    
    Dim endTime As Single
    endTime = timer + delayMillis * MillisToSeconds
    Do While timer < endTime
        DoEvents
        Dim hasBeenTriggered As Boolean
        If trigger And loggerErrorEvent Then hasBeenTriggered = hasBeenTriggered Or errorHappened
        If trigger And loggerLogEvent Then hasBeenTriggered = hasBeenTriggered Or logHappened
        If hasBeenTriggered Then Exit Do
    Loop
    errorHappened = False
    logHappened = False
End Sub
