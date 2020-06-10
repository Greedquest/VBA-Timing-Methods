VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetronomeExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Experiments")
'@IgnoreModule

Option Explicit
'@PredeclaredId

Private WithEvents ticker As OLDMetronome
Attribute ticker.VB_VarHelpID = -1
Private WithEvents ticker2 As OLDMetronome
Attribute ticker2.VB_VarHelpID = -1

Public Sub Init()
    ticker.Start
    ticker2.Start
End Sub

Private Sub Class_Initialize()
    Set ticker = OLDMetronomeCollection.Add(4000, True, -1)
    Set ticker2 = OLDMetronomeCollection.Add(1000, True, -1)
End Sub

Private Sub ticker_Tick()
    Application.Speech.Speak "Dum", True
End Sub

Private Sub ticker2_Tick()
    Application.Speech.Speak "Chaa", True
End Sub

