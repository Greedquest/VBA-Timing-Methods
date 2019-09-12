VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("FirstLevelAPI")

Option Explicit
'@PredeclaredId

Private Type tTimerData
    callbackWrapper As Object
    timerProc As LongPtr
End Type

Private this As tTimerData

Public Function Create(ByVal timerCallbackWrapper As Object) As TimerData
    With New TimerData
        Set .callbackWrapper = timerCallbackWrapper
        If TypeOf timerCallbackWrapper Is ManagedCallbackWrapper Then
            .timerProc = VBA.CLngPtr(AddressOf InternalTimerProcs.ManagedTimerCallbackInvoker)
        Else
            .timerProc = .callbackWrapperInterface.FunctionID
        End If
        Set Create = .Self
    End With
End Function

Friend Function Self() As TimerData
    Set Self = Me
End Function

Public Property Get callbackWrapperPointer() As LongPtr
    callbackWrapperPointer = ObjPtr(this.callbackWrapper)
End Property

Friend Property Get callbackWrapperInterface() As ICallbackWrapper
    Set callbackWrapperInterface = this.callbackWrapper
End Property

Public Property Set callbackWrapper(ByVal value As Object)
    Set this.callbackWrapper = value
End Property

Public Property Get id() As LongPtr 'alias
    id = Me.callbackWrapperPointer
End Property

Public Property Get timerProc() As LongPtr
    timerProc = this.timerProc
End Property

Friend Property Let timerProc(ByVal value As LongPtr)
    this.timerProc = value
End Property
