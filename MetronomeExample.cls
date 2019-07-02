VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetronomeExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Tests")

Option Explicit
'@PredeclaredId

Private WithEvents ticker As Metronome
Attribute ticker.VB_VarHelpID = -1

Sub Init()
    Set ticker = MetronomeCollection.Add(1000, True, -1)
End Sub


Private Sub ticker_Tick()
    Application.StatusBar = Now
End Sub
