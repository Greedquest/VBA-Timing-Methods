VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("FirstLevelAPI.Utils.Cache")

Option Explicit
'@PredeclaredId

Public Static Function loadObject(ByVal name As String, Optional ByVal default As Object) As Object

    ' @Reference("mscorlib.tlb 2.0")
    ' @Reference("Common Language Runtime Execution Engine")

    Dim host As New mscoree.CorRuntimeHost
    Dim domain As mscorlib.AppDomain
      
    If domain Is Nothing Then
        host.Start
        host.GetDefaultDomain domain
    End If

    If IsObject(domain.GetData(name)) Then
        Set loadObject = domain.GetData(name)
    Else
        Set loadObject = default
        domain.SetData name, default
    End If

End Function

