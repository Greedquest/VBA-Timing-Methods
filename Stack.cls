VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Simple LIFO Stack implementation which wraps a VBA.Collection"
Option Explicit
'@Folder("API.Utils.Stack")
'@ModuleDescription("Simple LIFO Stack implementation which wraps a VBA.Collection")

Private col As New Collection

Public Sub Push(ByVal item As Variant)
    col.Add item
End Sub

Public Function Pop() As Variant
    If col.Count > 0 Then
        LetSet Pop, col.item(col.Count)
        col.Remove col.Count
    End If
End Function

Public Function Peek() As Variant
    If col.Count > 0 Then
        LetSet Peek, col.item(col.Count)
    End If
End Function

Public Property Get Count() As Long
    Count = col.Count
End Property

Private Sub LetSet(ByRef variable As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If
End Sub

