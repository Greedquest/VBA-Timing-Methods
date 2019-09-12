VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FirstLevelAPI")
Option Explicit

Private Property Get timerExists(ByVal key As LongPtr) As Boolean
    On Error Resume Next                         'if there's a problem then the timerID is as good as unregistered anyway
    timerExists = this.timerIDs.Exists(key)
    On Error GoTo 0
End Property
