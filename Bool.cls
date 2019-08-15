VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Boolean object that can be passed byRef"
'@Exposed
'@Folder("FirstLevelAPI.Utils.Bool")
'@ModuleDescription("Boolean object that can be passed byRef")
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

Private Type tBool
    value As Boolean
End Type

Private this As tBool

'@Description("Truthy encapsulated value")
'@DefaultMember
Public Property Get value() As Boolean
    value = this.value
End Property

Public Property Let value(ByVal newVal As Boolean)
    this.value = newVal
End Property

'@Description("Create instance of Bool class from objPtr")
#If VBA7 Then
Public Function FromPtr(ByVal pData As LongPtr) As Bool
#Else
Public Function FromPtr(ByVal pData As Long) As Bool
#End If
Dim result As Object
CopyMemory result, pData, LenB(pData)
Set FromPtr = result
ZeroMemory result, LenB(pData)                   ' free up memory, equiv: CopyMemory result, 0&, LenB(pData)
End Function

'@Description("Class Constructor takes Boolean or Boolean-like values")
Public Function Create(ByVal initialValue As Variant) As Bool
    Dim result As New Bool
    result.value = CBool(initialValue)
    Set Create = result
End Function

