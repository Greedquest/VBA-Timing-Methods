VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TickerParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("FirstLevelAPI")
'@PredeclaredID
Option Explicit

Private Type tParams
    TickerIsRunning As Boolean
    data As Variant
End Type

Private this As tParams

Friend Function Create(Optional ByVal initialTickerState As Boolean = True, Optional ByVal data As Variant) As TickerParams
'NOTE only API needs to be able to create these so don't expose
    With New TickerParams
        .storeData IIf(IsMissing(data), Nothing, data)
        .TickerIsRunning = initialTickerState
        Set Create = .Self
    End With
End Function

Friend Function Self() As TickerParams
    Set Self = Me
End Function

Public Sub storeData(ByVal data As Variant)
    LetSet this.data, data
End Sub

Public Property Get storedData() As Variant
    LetSet storedData, this.data
End Property

Public Property Let TickerIsRunning(ByVal newState As Boolean)
    this.TickerIsRunning = newState
End Property

Public Property Get TickerIsRunning() As Boolean
    TickerIsRunning = this.TickerIsRunning
End Property

Private Sub Class_Initialize()
    Set this.data = Nothing
    this.TickerIsRunning = True
End Sub
