VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IResourceLocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("API.Utils.ResourceManager")
'@Exposed
Option Explicit

Public Function tryLocate(ByRef outResource As Variant, Optional ByVal resourceIdentifier As Variant) As Boolean
    'If resourceIdentifier is missing then locate default value
    If ResourceManager.isNothing(resourceIdentifier) Then
        Set outResource = Nothing
        tryLocate = True
    Else
        tryLocate = False
    End If
End Function
