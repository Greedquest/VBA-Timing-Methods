VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnmanagedCallbackWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("FirstLevelAPI.Utils.Wrappers")
'@PredeclaredID
'@Exposed
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

Private Type tUnmanagedWrapper
    callbackFunction As LongPtr
    data As Variant
End Type

Private this As tUnmanagedWrapper

Friend Function Create(ByVal callbackFunction As LongPtr, Optional ByVal data As Variant) As UnmanagedCallbackWrapper
    'NOTE only API needs to be able to create these so don't expose
    With New UnmanagedCallbackWrapper
        .storeData IIf(IsMissing(data), Nothing, data)
        .Callback = callbackFunction
        Set Create = .Self
    End With
End Function

Friend Function Self() As UnmanagedCallbackWrapper
    Set Self = Me
End Function

Public Sub storeData(ByVal data As Variant)
    LetSet this.data, data
End Sub

Public Property Get storedData() As Variant
    LetSet storedData, this.data
End Property

Friend Property Let Callback(ByVal value As LongPtr)
    this.callbackFunction = value
End Property

#If VBA7 Then
Public Function FromPtr(ByVal pData As LongPtr) As UnmanagedCallbackWrapper
#Else
Public Function FromPtr(ByVal pData As Long) As UnmanagedCallbackWrapper
#End If
Dim result As Object
Dim timerStuff As Dictionary
Set timerStuff = Cache.loadObject("TickerAPI" & ".timerIDs", New Dictionary)
Debug.Print IIf(timerStuff.Count > 0, "Some timer about", "It's not in the cache");
If timerStuff.Count = 1 Then
    Debug.Print " - "; timerStuff.Keys(0), printf("({0})", timerStuff.item(timerStuff(timerStuff.Keys(0))).debugName)
Else
    Debug.Print
End If
CopyMemory result, pData, LenB(pData)
Set FromPtr = result                             'don't copy directly as then reference count won't be managed (I think)
ZeroMemory result, LenB(pData)                   ' free up memory, equiv: CopyMemory result, 0&, LenB(pData)
End Function

Private Sub Class_Initialize()
    Set this.data = Nothing
    this.callbackFunction = 0
End Sub

Private Property Get ICallbackWrapper_Callback() As Variant
    ICallbackWrapper_Callback = this.callbackFunction
End Property

Private Property Get ICallbackWrapper_CallbackType() As CallbackType
    ICallbackWrapper_CallbackType = cbRawPointer
End Property

Friend Static Property Get debugName() As String
    Dim name As String
    If name = vbNullString Then
        Dim i As Long
        For i = 1 To 7
            name = name & Chr$(WorksheetFunction.RandBetween(65, 90))
        Next
    End If
    debugName = name
End Property

Public Property Get timerID() As LongPtr
    timerID = ObjPtr(Me)
End Property
