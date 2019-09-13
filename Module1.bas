Attribute VB_Name = "Module1"

Option Explicit

Sub t()
Dim a As New TimerRepository
a.Add UnmanagedCallbackWrapper.Create(AddressOf t, "boo")
a.Add ManagedCallbackWrapper.Create(New HelloWorldTimerProc, "ya")
Dim arr() As Variant
arr = a.ToArray
End Sub
