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

'@DefaultMember
Public Function Item(ByVal timerID As LongPtr) As TimerData
Attribute Item.VB_UserMemId = 0
    Set Item = this.TimerData.Item(timerID)
End Function

Public Function Add(ByVal callbackWrapper As Object) As TimerData
    Dim newData As TimerData
    Set newData = TimerData.Create(callbackWrapper)
    this.TimerData.Add newData.ID, newData
    Set Add = newData
End Function

Public Sub Remove(ByVal timerInfo As TimerData)
    this.TimerData.Remove timerInfo.ID
End Sub

Public Sub Clear()
    this.TimerData.RemoveAll
End Sub

Public Function ToArray() As Variant
    ToArray = this.TimerData.Items
End Function

Public Property Get Exists(ByVal timerID As LongPtr) As Boolean
    On Error Resume Next                         'if there's a problem then the timerID is as good as unregistered anyway
    Exists = this.TimerData.Exists(timerID)
    On Error GoTo 0
End Property

Public Function FilterByFunctionID(ByVal funcID As LongPtr) As Collection
    Dim matches As New Collection
    Dim data As TimerData
    For Each data In this.TimerData
        If data.callbackWrapperInterface.FunctionID = funcID Then
            matches.Add data
        End If
    Next data
    Set FilterByFunctionID = matches
End Function

Public Function FilterByFunctionName(ByVal funcName As String) As Collection
    Dim matches As New Collection
    Dim data As TimerData
    For Each data In this.TimerData
        If data.callbackWrapperInterface.FunctionName = funcName Then
            matches.Add data
        End If
    Next data
    Set FilterByFunctionName = matches
End Function

