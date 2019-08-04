VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersistentDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("API.Utils.PersistentDict")
'@PredeclaredID
Option Explicit

Private dict As Scripting.Dictionary

Public Function Create(ByVal name As String) As PersistentDict
    With New PersistentDict
        .setPersistentDictionary name
        Set Create = .Self
    End With
End Function

Public Property Get Data() As Scripting.Dictionary
    Set Data = dict
End Property


Friend Property Get Self() As PersistentDict
    Set Self = Me
End Property

Friend Sub setPersistentDictionary(ByVal name As String)
    ' References:
    '  mscorlib.dll
    '  Common Language Runtime Execution Engine

    If dict Is Nothing Then
      Dim host As New mscoree.CorRuntimeHost
      Dim domain As mscorlib.AppDomain
      host.Start
      host.GetDefaultDomain domain

      If IsObject(domain.GetData(name)) Then
        Set dict = domain.GetData(name)
      Else
        Set dict = New Scripting.Dictionary
        domain.SetData name, dict
      End If
    End If

End Sub
