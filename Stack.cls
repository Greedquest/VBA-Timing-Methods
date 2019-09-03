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
'@Folder("Old.FirstLevelAPI.Utils.Stack")
'@ModuleDescription("Simple LIFO Stack implementation which wraps a VBA.Collection")

Private col As New Collection

Public Sub Push(ByVal Item As Variant)
    col.Add Item
End Sub

Public Function Pop() As Variant
    If col.Count > 0 Then
        LetSet Pop, col.Item(col.Count)
        col.Remove col.Count
    End If
End Function

Public Function Peek() As Variant
    If col.Count > 0 Then
        LetSet Peek, col.Item(col.Count)
    End If
End Function

Public Property Get Count() As Long
    Count = col.Count
End Property
