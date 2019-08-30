Attribute VB_Name = "ProjectUtils"
'@Folder("Common")
Option Explicit
Option Private Module

Public Const INFINITE_DELAY As Long = &H7FFFFFFF

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub LetSet(ByRef variable As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If
End Sub

Public Sub throwDllError(ByVal ErrorNumber As Long, Optional ByVal onZeroText As String = "DLL error = 0, i.e. no error")
    If ErrorNumber = 0 Then
        Err.Raise 5, Description:=onZeroText
    Else
        Err.Raise ErrorNumber, Description:=GetSystemErrorMessageText(ErrorNumber)
    End If
End Sub

Public Sub logError(ByVal Source As String, ByVal errNum As Long, ByVal errDescription As String)
    If Not LogManager.IsEnabled(ErrorLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("Timing-E", ErrorLevel)
    End If
    LogManager.log ErrorLevel, Toolbox.Strings.Format("{0} raised an error: #{1} - {2}", Source, errNum, errDescription)
End Sub

Public Sub log(ByVal loggerLevel As LogLevel, ByVal Source As String, ByVal message As String)
    If Not LogManager.IsEnabled(TraceLevel) Then 'check a logger is registered
        LogManager.Register DebugLogger.Create("Timing", TraceLevel)
    End If
    LogManager.log loggerLevel, Toolbox.Strings.Format("{0} - {1}", Source, message)
End Sub

Sub t()
    Dim i As Long
    For i = 1 To 10000
        On Error Resume Next
        Err.Raise i
        If Err.Description <> "Application-defined or object-defined error" Then Debug.Print i, Err.Description
    Next i
    
    For i = 1 To 10000
        On Error Resume Next
        Err.Raise vbObjectError + i
        If Err.Description <> "Automation error" Then Debug.Print i, Replace(Replace(Err.Description, vbCrLf, vbNullString), vbLf, ": ")
    Next i
End Sub
