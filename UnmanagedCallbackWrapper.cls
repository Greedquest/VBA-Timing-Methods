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

Private Type tUnmanagedWrapper
    callbackFunction As LongPtr
    data As Variant
    Name As String
End Type

Private this As tUnmanagedWrapper

Private Sub Class_Initialize()
    Set this.data = Nothing
    this.callbackFunction = 0
    'TODO allow custom name
    this.Name = WinAPI.GetGUID                   'something unique to the function; could be the ptr but that might be reallocated
End Sub

Friend Function Create(ByVal callbackFunction As LongPtr, Optional ByVal data As Variant) As UnmanagedCallbackWrapper
    'NOTE only API needs to be able to create these so don't expose
    With New UnmanagedCallbackWrapper
        .storeData IIf(IsMissing(data), Nothing, data)
        .callBack = callbackFunction
        Set Create = .Self
    End With
End Function

Friend Function Self() As UnmanagedCallbackWrapper
    Set Self = Me
End Function

Friend Property Let callBack(ByVal value As LongPtr)
    this.callbackFunction = value
End Property

Public Sub storeData(ByVal data As Variant)
    LetSet this.data, data
End Sub

Public Property Get userData() As Variant
    LetSet userData, this.data
End Property

Public Property Get timerID() As LongPtr
    timerID = ObjPtr(Me)
End Property

Private Property Get ICallbackWrapper_FunctionID() As LongPtr
    ICallbackWrapper_FunctionID = this.callbackFunction
End Property

Private Property Get ICallbackWrapper_FunctionName() As String
    ICallbackWrapper_FunctionName = this.Name
End Property

'for testing
Friend Property Get debugName() As String
    debugName = this.Name
End Property
