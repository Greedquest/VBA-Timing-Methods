VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fakeDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dict As New Scripting.Dictionary

Public Property Get Item(ByVal key As String) As String
    Item = dict.Item(key)
End Property

Public Property Let Item(ByVal key As String, ByVal value As String)
    dict.Item(key) = value
End Property

Private Sub Class_Terminate()
    Debug.Print "I am made dead"
End Sub

