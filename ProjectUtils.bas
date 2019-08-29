Attribute VB_Name = "ProjectUtils"
'@Folder("Common")
Option Explicit
Option Private Module

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub LetSet(ByRef variable As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If
End Sub

Public Sub raiseDllError(ByVal ErrorNumber As Long, Optional ByVal Source As String = "raiseDllError")
    Err.Description = GetSystemErrorMessageText(ErrorNumber)
    logError Source, ErrorNumber, Err.Description
    Err.Raise ErrorNumber
End Sub

Public Sub logError(ByVal Source As String, ByVal errNum As Long, ByVal errDescription As String)
    If Not LogManager.IsEnabled(ErrorLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("Timing", ErrorLevel)
    End If
    LogManager.log ErrorLevel, Toolbox.Strings.Format("{0} raised an error: #{1} - {2}", Source, errNum, errDescription)
End Sub

Public Sub log(ByVal loggerLevel As LogLevel, ByVal Source As String, ByVal message As String)
    If Not LogManager.IsEnabled(TraceLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("Timing" & TraceLevel, TraceLevel)
    End If
    LogManager.log loggerLevel, Toolbox.Strings.Format("{0} - {1}", Source, message)
End Sub
