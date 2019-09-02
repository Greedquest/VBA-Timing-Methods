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
    name As String
End Type

Private this As tUnmanagedWrapper

Private Sub Class_Initialize()
    Set this.data = Nothing
    this.callbackFunction = 0
    this.name = WinAPI.GetGUID 'something unique to the function; could be the ptr but that might be reallocated
End Sub

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

Public Property Get userData() As Variant
    LetSet userData, this.data
End Property

Friend Property Let Callback(ByVal value As LongPtr)
    this.callbackFunction = value
End Property

Friend Static Property Get debugName() As String
    debugName = this.name
End Property

Public Property Get timerID() As LongPtr
    timerID = ObjPtr(Me)
End Property

Private Property Get ICallbackWrapper_CallbackID() As LongPtr
    ICallbackWrapper_CallbackID = this.callbackFunction
End Property

Private Property Get ICallbackWrapper_CallbackName() As String
    ICallbackWrapper_CallbackName = this.name
End Property

