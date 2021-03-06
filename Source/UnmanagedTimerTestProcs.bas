Attribute VB_Name = "UnmanagedTimerTestProcs"
'@Folder("Tests")
Option Explicit

Private Type tTestData
    testLog As New testLog
End Type

Private this As tTestData

'@Ignore ParameterNotUsed: callbacks need to have this signature regardless
Public Sub UnmanagedTimerTestProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal callbackParams As UnmanagedCallbackWrapper, ByVal tickCount As Long)
    Const loggerSourceName As String = "timerProc"
    On Error GoTo CleanFail
    this.testLog.logCall callbackParams.timerID, callbackParams.userData
    
CleanExit:
    Exit Sub
    
CleanFail:
    this.testLog.logError ObjPtr(callbackParams), Err.Number, Err.Description, loggerSourceName
    Resume CleanExit
End Sub

Public Property Get testLog() As testLog
    Set testLog = this.testLog
End Property

Public Sub clearLog()
    Set this.testLog = Nothing
End Sub

