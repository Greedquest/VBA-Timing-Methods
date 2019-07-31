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

'@ModuleDescription("Example dictionary resource; the resource provided is a key in the encapsulated dictionary. No locator is provided so obtain resource fails with the same resourceIdentifier.)

Implements IResourceCreator
Implements IResourceDestroyer

Public encapsulated As New Scripting.Dictionary

Private Function IResourceCreator_tryCreate(Optional ByVal resourceIdentifier As Variant, Optional ByVal resourceArgs As Variant, Optional ByRef outResource As Variant) As Boolean
    If Not IsMissing(resourceIdentifier) Then
        ResourceManager.LetSet outResource, resourceIdentifier
        On Error Resume Next
        encapsulated.Add resourceIdentifier, resourceArgs
        If Err.Number = 0 Then IResourceCreator_tryCreate = True
        On Error GoTo 0
    End If
End Function


Private Function IResourceDestroyer_tryDestroy(ByVal resource As Variant) As Boolean
    On Error Resume Next
    encapsulated.Remove resource
    If Err.Number = 0 Then IResourceDestroyer_tryDestroy = True
    On Error GoTo 0
End Function
