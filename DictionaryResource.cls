VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("API.Utils.ResourceManager.Tests")
Option Explicit

Implements IResourceCreator
Implements IResourceDestroyer

Private encapsulates As New Scripting.Dictionary

Private Function IResourceCreator_tryCreate(ByRef outResource As Variant, Optional ByVal resourceIdentifier As Variant, Optional ByVal resourceArgs As Variant) As Boolean
    If IsMissing(resourceIdentifier) Then
        IResourceCreator_tryCreate = False
    Else
    End If
End Function


Private Function IResourceDestroyer_tryDestroy(ByVal resource As Variant) As Boolean
    Err.Raise 5 'TODO implement interface member
End Function

