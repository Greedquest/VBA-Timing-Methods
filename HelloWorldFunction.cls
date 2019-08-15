VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelloWorldFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Tests.Callbacks")

Implements ITimerProc

Private Function ITimerProc_Exec() As Boolean
    Debug.Print "Hey there matey"
    ITimerProc_Exec = True
End Function
