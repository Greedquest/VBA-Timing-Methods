VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITimerProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("FirstLevelAPI.Utils.Functions")
'@PredeclaredID
Option Explicit

'NOTE RtlCopyMemory would be faster (as source and dest won't overlap) but is not exposed to VBA
''@Description("API: Destination and Source can be byVal pointers or byRef variables, length is LenB(dataType)")
#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)
#End If

#If VBA7 Then
    Friend Function FromPtr(ByVal pData As LongPtr) As ITimerProc
#Else
    Friend Function FromPtr(ByVal pData As Long) As ITimerProc
#End If
    Dim result As Object
    CopyMemory result, pData, LenB(pData)
    Set FromPtr = result 'don't copy directly as then reference count won't be managed (I think)
    ZeroMemory result, LenB(pData)                   ' free up memory, equiv: CopyMemory result, 0&, LenB(pData)
End Function

'Interface
Public Function Exec() As Boolean
    Err.Raise 5 'not implemented
End Function
