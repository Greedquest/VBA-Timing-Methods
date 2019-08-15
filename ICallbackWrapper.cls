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

Public Enum CallbackType
    cbRawPointer
    cbCommandObject
End Enum

Public Property Get Callback() As Variant
End Property

'Public Sub setCallback(ByVal value As Variant)
'End Sub

Public Property Get CallbackType() As CallbackType
End Property
