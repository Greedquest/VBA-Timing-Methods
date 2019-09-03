Attribute VB_Name = "InternalTimerProcs"
'@Folder("FirstLevelAPI.Utils")
Option Explicit
Option Private Module

Private Const killOnError As Boolean = True      'TODO make this configurable

'@Description("TIMERPROC callback for ManagedCallbacks which executes the callback function within error guards")
Public Sub ManagedTimerCallbackInvoker(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerParams As ManagedCallbackWrapper, ByVal tickCount As Long)
    Const loggerSourceName As String = "ManagedTimerCallbackInvoker"
    
    'NOTE could check message and ObjPtr(timerparams) to ensure this is a valid managedTimer caller
    On Error Resume Next
    timerParams.Callback.Exec timerParams.timerID, timerParams.storedData, tickCount
    
    Dim errNum As Long
    Dim errDescription As String
    errNum = Err.Number                          'changing the error policy will wipe these, so cache them
    errDescription = Err.Description
    
    'Log any error the callback may have raised, kill it if necessary
    On Error GoTo cleanFail                      'this procedure cannot raise errors or we'll crash
    If errNum <> 0 Then
        logError timerParams.CallbackWrapper.FunctionName & ".Exec", errNum, errDescription
        If killOnError Then
            On Error GoTo cleanFail
            TickerAPI.KillTimerByID timerParams.timerID
        End If
    End If
    
cleanExit:
    Exit Sub
    
cleanFail:
    logError loggerSourceName, Err.Number, Err.Description
    Resume cleanExit
End Sub
