VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "testLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Common")
Option Explicit

Private Type testLogData
    callLog As New Scripting.Dictionary          '{callerID:[data1,data2...]}
    errorLog As New Scripting.Dictionary         '{callerID:[(errNum1,errInfo1), (errNum2,errInfo2),...]}
End Type

Public Enum loggerEventType
    loggerErrorEvent = 2 ^ 0
    loggerLogEvent = 2 ^ 1
    loggerAllEvents = -1
End Enum

Private this As testLogData

Public Sub logError(ByVal callerID As LongPtr, ByVal errNum As Long, ByVal errDescription As String, Optional ByVal Source As String = "TimerProc")
    If Not this.errorLog.Exists(callerID) Then
        this.errorLog.Add callerID, New Collection
    End If
    Dim errorSet As Collection
    Set errorSet = this.errorLog.Item(callerID)
    errorSet.Add Array(errNum, Toolbox.Strings.Format("{0} caused an error in {3}: #{1} - {2}", callerID, errNum, errDescription, Source))
End Sub

Public Sub logCall(ByVal callerID As LongPtr, ByVal data As Variant)
    If Not this.callLog.Exists(callerID) Then
        this.callLog.Add callerID, New Collection
    End If
    Dim callSet As Collection
    Set callSet = this.callLog.Item(callerID)
    callSet.Add data
End Sub

Public Property Get errorLog() As Dictionary
    Set errorLog = this.errorLog
End Property

Public Property Get callLog() As Dictionary
    Set callLog = this.callLog
End Property

Public Property Get errorCount(Optional ByVal callerID As LongPtr = 0) As Long
    Dim errCount As Long
    If callerID = 0 Then                         'total
        Dim errorSet As Variant
        For Each errorSet In this.errorLog.Items
            errCount = errCount + errorSet.Count
        Next errorSet
    Else
        On Error Resume Next
        errCount = this.errorLog(callerID).Count
        On Error GoTo 0
    End If
    errorCount = errCount
End Property

Public Property Get callCount(Optional ByVal callerID As LongPtr = 0) As Long
    If callerID = 0 Then                         'total
        Dim runningTot As Long
        Dim callSet As Variant
        For Each callSet In this.callLog.Items
            runningTot = runningTot + callSet.Count
        Next callSet
        callCount = runningTot
    Else
        On Error Resume Next
        callCount = this.callLog(callerID).Count
        On Error GoTo 0
    End If
End Property

'@Description("Waits until next instance of the events given by trigger, up to a maximum delay. Returns False if never triggered")
Public Function waitUntilTrigger(Optional ByVal maxDelayMillis As Long = defaultTimerDelay, Optional ByVal callerID As LongPtr = 0, Optional ByVal trigger As loggerEventType = loggerAllEvents) As Boolean
Attribute waitUntilTrigger.VB_Description = "Waits until next instance of the events given by trigger, up to a maximum delay. Returns False if never triggered"
    
    Dim endTime As Single
    endTime = timer + maxDelayMillis * MillisToSeconds
    
    Dim lastErrCount As Long
    Dim lastCallCount As Long
    
    Dim hasBeenTriggered As Boolean
    
    Do While timer < endTime
        updateCounts lastErrCount, lastCallCount, callerID
        DoEvents
        'see what constitutes a trigger
        If trigger And loggerLogEvent Then
            hasBeenTriggered = callCount(callerID) > lastCallCount Or hasBeenTriggered
        End If
        If trigger And loggerErrorEvent Then
            hasBeenTriggered = errorCount(callerID) > lastErrCount Or hasBeenTriggered
        End If
        If hasBeenTriggered Then Exit Do
    Loop
    
    waitUntilTrigger = hasBeenTriggered
End Function

Private Sub updateCounts(ByRef outErrCount As Long, ByRef outCallCount As Long, ByVal callerID As LongPtr)
    outErrCount = errorCount(callerID)
    outCallCount = callCount(callerID)
End Sub
