VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResourceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("API.Utils.ResourceManager")
'@Exposed
'@PredeclaredId
Option Explicit

Private Type tResourceManager
    Creator As IResourceCreator
    Destroyer As IResourceDestroyer
    Locator As IResourceLocator
End Type

Public Enum ResourceManagerError
    ObtainResourceError = vbObjectError + 1
    DestroyResourceError
End Enum

Private this As tResourceManager

Private Sub Class_Initialize()
    'set defaults so that the fresh resourceManager isn't totally invalid
    Set Me.Creator = New IResourceCreator
    Set Me.Destroyer = New IResourceDestroyer
    Set Me.Locator = New IResourceLocator
End Sub

'''Factory & Access Methods
'NOTE use Friend modifier for anything used to build the classes intenally, Public for exposing anything provided by caller code
Public Function Create(ByVal resourceCreator As IResourceCreator, ByVal resourceDestroyer As IResourceDestroyer, Optional ByVal resourceLocator As IResourceLocator) As ResourceManager
    With New ResourceManager
        Set .Creator = resourceCreator
        Set .Destroyer = resourceDestroyer
        If Not isNothing(resourceLocator) Then Set .Locator = resourceLocator
        Set Create = .Self
    End With
End Function

Friend Function Self() As ResourceManager
    Set Self = Me
End Function

Friend Property Set Creator(ByVal value As IResourceCreator)
    Set this.Creator = value
End Property

Public Property Get Creator() As IResourceCreator
    Set Creator = this.Creator
End Property

Friend Property Set Destroyer(ByVal value As IResourceDestroyer)
    Set this.Destroyer = value
End Property

Public Property Get Destroyer() As IResourceDestroyer
    Set Destroyer = this.Destroyer
End Property

Friend Property Set Locator(ByVal value As IResourceLocator)
    Set this.Locator = value
End Property

Public Property Get Locator() As IResourceLocator
    Set Locator = this.Locator
End Property

'''Util methods - put here as the Friend modifier can't be used in normal modules
Friend Function isNothing(ByVal arg As Variant) As Boolean
    If IsObject(arg) Then
        isNothing = arg Is Nothing
    Else
        isNothing = False
    End If
End Function

Friend Function LetSet(ByRef variable As Variant, ByVal value As Variant) As Boolean
    If IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If
End Function

'''Public interface

'@Description("Attempt to destroy specified resource, or if missing then find and destroy any resources. Returns True if resource destroyed, False if none found, Raises Error if unable to destroy.")
Public Function ReleaseResource(Optional ByVal resource As Variant) As Boolean
Attribute ReleaseResource.VB_Description = "Attempt to destroy specified resource, or if missing then find and destroy any resources. Returns True if resource destroyed, False if none found, Raises Error if unable to destroy."
    If IsMissing(resource) Then
        Dim foundResource As Variant
        If this.Locator.tryLocate(foundResource) Then
            If Not this.Destroyer.tryDestroy(resource) Then
                Err.Raise DestroyResourceError, Description:="Resource located but could not be destroyed"
                Exit Function
            End If
        Else
            'NO-OP, nothing to destroy
            ReleaseResource = False
        End If
    Else
        If Not this.Destroyer.tryDestroy(resource) Then
            Err.Raise DestroyResourceError, Description:="Unable to destroy specified resource"
            Exit Function
        End If
        ReleaseResource = True
    End If
End Function

'@Description("First tries to locate matching resources (ignoring resourceArgs), if that fails then attempts to create one. Returns the resource if found, RaisesError if unable to find or create one.")
Public Function ObtainResource(ByVal resourceIdentifier As Variant, Optional ByVal resourceArgs As Variant) As Variant
Attribute ObtainResource.VB_Description = "First tries to locate matching resources (ignoring resourceArgs), if that fails then attempts to create one. Returns the resource if found, RaisesError if unable to find or create one."
    Dim resource As Variant
    If Not this.Locator.tryLocate(resource, resourceIdentifier) Then
        If Not this.Creator.tryCreate(resource, resourceIdentifier, resourceArgs) Then
            Err.Raise ObtainResourceError, Description:="Couldn't locate or create specified resource"
            Exit Function
        End If
    End If
    LetSet ObtainResource, resource
End Function
