Attribute VB_Name = "InternalTimerProcs"
'@Folder("FirstLevelAPI.Utils")
Option Explicit
Option Private Module

'@Description("TIMERPROC callback for ManagedCallbacks which executes the callback function within error guards")
Public Sub ManagedTimerCallbackInvoker(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerParams As ManagedCallbackWrapper, ByVal tickCount As Long)
    'NOTE could check message and ObjPtr(timerparams) to ensyure this is a valid managedTimer caller
    On Error GoTo cleanFail
    timerParams.Callback.Exec timerParams.timerID, timerParams.storedData, tickCount
    
cleanExit:
    Exit Sub
    
cleanFail:
     logError timerParams.CallbackWrapper.FunctionName & ".Exec", Err.Number, Err.Description
     Resume cleanExit
End Sub
