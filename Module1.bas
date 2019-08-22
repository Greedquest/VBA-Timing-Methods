Attribute VB_Name = "Module1"
Option Explicit

Sub testcasting()

Debug.Print "Initialize"
Dim a As Object
Dim b As ICallbackWrapper
Dim c As Object
Debug.Print Toolbox.Strings.Format("\ta: {0} - {1}\n\tb: {2} - {3}\n\tc: {4} - {5}", ObjPtr(a), TypeName(a), ObjPtr(b), TypeName(b), ObjPtr(c), TypeName(c))

Debug.Print "DownCast"
Set a = New UnmanagedCallbackWrapper
Set b = a
Debug.Print Toolbox.Strings.Format("\ta: {0} - {1}\n\tb: {2} - {3}\n\tc: {4} - {5}", ObjPtr(a), TypeName(a), ObjPtr(b), TypeName(b), ObjPtr(c), TypeName(c))

Debug.Print "Upcast"
Set c = b
Debug.Print Toolbox.Strings.Format("\ta: {0} - {1}\n\tb: {2} - {3}\n\tc: {4} - {5}", ObjPtr(a), TypeName(a), ObjPtr(b), TypeName(b), ObjPtr(c), TypeName(c))

End Sub


Sub makeIll()
    TickerAPI.makeBadness
End Sub
