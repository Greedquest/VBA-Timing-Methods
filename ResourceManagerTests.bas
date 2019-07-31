Attribute VB_Name = "ResourceManagerTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("API.Utils.ResourceManager.Tests")

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set ResourceManager = New ResourceManager 'reset default instance?
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub ObtainResourceWithDefaultLocatorAndCreatorRaisesError()
    Const ExpectedError As Long = ResourceManagerError.ObtainResourceError
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    ResourceManager.ObtainResource "Joe", "bloggs" 'don't even bother with return value

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Uncategorized")
Private Sub ReleaseSpecifiedResourceWithDefaultDestroyerRaisesError()
    Const ExpectedError As Long = ResourceManagerError.DestroyResourceError
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    ResourceManager.ReleaseResource "Joe"

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Uncategorized")
Private Sub ReleaseUnspecifiedResourceWithDefaultLocatorAndDestroyerReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim result As Boolean
    
    'Act:
    result = ResourceManager.ReleaseResource()
    
    'Assert:
    Assert.IsFalse result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

