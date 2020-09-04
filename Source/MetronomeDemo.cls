VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetronomeDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Experiments")
'@IgnoreModule

Option Explicit
'@PredeclaredId

Private WithEvents ticker As Metronome
Attribute ticker.VB_VarHelpID = -1
Private WithEvents ticker2 As Metronome
Attribute ticker2.VB_VarHelpID = -1

Public Sub Init()
    ticker.Start
    ticker2.Start
End Sub

Private Sub Class_Initialize()
    Set ticker = Metronome.Create(1000, delayedStart:=True)
    Set ticker2 = Metronome.Create(700, delayedStart:=True)
End Sub

Private Sub ticker_Tick()
    Debug.Print "Tick"
End Sub

Private Sub ticker2_Tick()
    Debug.Print "Tock"
End Sub

