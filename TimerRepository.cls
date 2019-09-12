VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FirstLevelAPI")
Option Explicit

Private Type repositoryData
    TimerData As New Scripting.Dictionary '{id:TimerData}
End Type

Private this As repositoryData

Private Property Get timerExists(ByVal key As LongPtr) As Boolean
    On Error Resume Next                         'if there's a problem then the timerID is as good as unregistered anyway
    timerExists = this.timerIDs.Exists(key)
    On Error GoTo 0
End Property

Public Function getTimerByID(ByVal id As LongPtr) As TimerData
    
End Function

Public Function getTimerArrayByFunctionID(ByVal FunctionID As LongPtr) As TimerData()
    
End Function

Public Function getTimerArrayByFunctionName(ByVal FunctionName As String) As TimerData()

End Function

Public Function Add(ByVal callbackWrapper As Object) As TimerData
'TODO log error
    Dim newData As TimerData
    Set newData = TimerData.Create(callbackWrapper)
    this.TimerData.Add data.id, data
    Set Add = newData
    
End Function

Public Function Exists(ByVal timerID As LongPtr) As Boolean
    Exists = this.TimerData.Exists(timerID)
End Function
