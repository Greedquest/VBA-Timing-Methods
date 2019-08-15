VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManagedCallbackWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("FirstLevelAPI.Utils.Wrappers")
'@PredeclaredID
Option Explicit

Implements ICallbackWrapper

'NOTE RtlCopyMemory would be faster (as source and dest won't overlap) but is not exposed to VBA
''@Description("API: Destination and Source can be byVal pointers or byRef variables, length is LenB(dataType)")
#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)
#End If

Private Type tManagedWrapper
    callbackFunction As ITimerProc
    data As Variant
End Type

Private this As tManagedWrapper

Friend Function Create(ByVal callbackFunction As ITimerProc, Optional ByVal data As Variant) As ManagedCallbackWrapper
    'NOTE only API needs to be able to create these so don't expose
    With New ManagedCallbackWrapper
        .storeData IIf(IsMissing(data), Nothing, data)
        Set Create = .Self
    End With
End Function

Public Function Self() As ManagedCallbackWrapper
    Set Self = Me
End Function

Public Sub storeData(ByVal data As Variant)
    LetSet this.data, data
End Sub

Public Property Get storedData() As Variant
    LetSet storedData, this.data
End Property

Public Property Set Callback(ByVal obj As ITimerProc)
    Set this.callbackFunction = obj
End Property

#If VBA7 Then
Public Function FromPtr(ByVal pData As LongPtr) As ManagedCallbackWrapper
#Else
Public Function FromPtr(ByVal pData As Long) As TimerProcWrapper
#End If
Dim result As Object
CopyMemory result, pData, LenB(pData)
Set FromPtr = result                             'don't copy directly as then reference count won't be managed (I think)
ZeroMemory result, LenB(pData)                   ' free up memory, equiv: CopyMemory result, 0&, LenB(pData)
End Function

Private Property Get ICallbackWrapper_Callback() As Variant
    Set ICallbackWrapper_Callback = this.callbackFunction
End Property

Private Property Get ICallbackWrapper_CallbackType() As CallbackType
    ICallbackWrapper_CallbackType = cbCommandObject
End Property

Private Sub Class_Initialize()
    Set this.data = Nothing
    Set this.callbackFunction = New ITimerProc
End Sub

