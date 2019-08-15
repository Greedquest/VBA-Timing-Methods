Attribute VB_Name = "ProjectUtils"
'@Folder("ProjectUtils")
Option Explicit
Option Private Module

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub LetSet(ByRef variable As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If
End Sub

