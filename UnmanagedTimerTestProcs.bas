Attribute VB_Name = "UnmanagedTimerTestProcs"
'@Folder("Tests")
Option Explicit

Private Type tTestData
    testLog As New testLog
End Type

Private this As tTestData

Public Sub UnmanagedTimerTestProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal callbackParams As UnmanagedCallbackWrapper, ByVal tickCount As Long)
    Const loggerSourceName As String = "timerProc"
    On Error GoTo cleanFail
    this.testLog.logCall callbackParams.timerID, callbackParams.userData
    
cleanExit:
    Exit Sub
    
cleanFail:
    this.testLog.logError ObjPtr(callbackParams), Err.Number, Err.Description, loggerSourceName
    Resume cleanExit
End Sub

Public Property Get testLog() As testLog
    Set testLog = this.testLog
End Property

Public Sub clearLog()
    Set this.testLog = Nothing
End Sub

Sub t()
    clearLog
    
    Dim runner As New scheduler
    runner.doEventsWait this.testLog
    Debug.Print printf("callCount: {0} errCount: {1}", this.testLog.callCount, this.testLog.errorCount)
    
    Dim timerID As LongPtr
    timerID = TickerAPI.StartUnmanagedTimer(AddressOf UnmanagedTimerTestProc, True)
    
    'Dim runner As New scheduler
    runner.doEventsWait this.testLog
    Debug.Print printf("callCount: {0} errCount: {1}", this.testLog.callCount, this.testLog.errorCount)
    
    TickerAPI.KillTimerByID timerID
End Sub