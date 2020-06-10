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

Private Type tManagedWrapper
    callbackFunction As ITimerProc
    data As Variant
End Type

Private this As tManagedWrapper

Private Sub Class_Initialize()
    Set this.data = Nothing
    Set this.callbackFunction = New ITimerProc
End Sub

Public Function Create(ByVal callbackFunction As ITimerProc, Optional ByVal data As Variant) As ManagedCallbackWrapper
    'NOTE only API needs to be able to create these so don't expose
    With New ManagedCallbackWrapper
        .storeData data
        Set .callBack = callbackFunction
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ManagedCallbackWrapper
    Set Self = Me
End Property

Public Property Set callBack(ByVal obj As ITimerProc)
    Set this.callbackFunction = obj
End Property

Public Property Get callBack() As ITimerProc
    Set callBack = this.callbackFunction
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
    ICallbackWrapper_FunctionID = ObjPtr(this.callbackFunction)
End Property

Private Property Get ICallbackWrapper_FunctionName() As String
    ICallbackWrapper_FunctionName = TypeName(this.callbackFunction)
End Property

Public Property Get callbackWrapper() As ICallbackWrapper 'just return the interface; makes it easier to work with
    Set callbackWrapper = Me
End Property

