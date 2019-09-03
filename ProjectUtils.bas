Attribute VB_Name = "ProjectUtils"
'@Folder("Common")
Option Explicit
Option Private Module

Public Const INFINITE_DELAY As Long = &H7FFFFFFF

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal length As Long)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)
#End If

#If VBA7 Then

Public Function FromPtr(ByVal pData As LongPtr) As Object
#Else
Public Function FromPtr(ByVal pData As Long) As Object
#End If
Dim result As Object
CopyMemory result, pData, LenB(pData)
Set FromPtr = result                             'don't copy directly as then reference count won't be managed (I think)
ZeroMemory result, LenB(pData)                   ' free up memory, equiv: CopyMemory result, 0&, LenB(pData)
End Function

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

