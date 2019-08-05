VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IResourceDestroyer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Old.ResourceManager")
'@Exposed
Option Explicit

Public Function tryDestroy(ByVal resource As Variant) As Boolean
    tryDestroy = ResourceManager.isNothing(resource) 'only fails if there was something to destroy!
End Function

