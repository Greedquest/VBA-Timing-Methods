VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICallbackWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("FirstLevelAPI.Utils.Wrappers")
'@Exposed
Option Explicit

Public Property Get CallbackID() As LongPtr
End Property

Public Property Get CallbackName() As String
End Property
