Attribute VB_Name = "Module1"
Option Explicit

Sub ttt()
Dim a As New Dictionary
Dim b As New Collection
Set a.item(b) = b
Debug.Print a.Exists(b)
Debug.Print a.Exists(ObjPtr(b))
End Sub
