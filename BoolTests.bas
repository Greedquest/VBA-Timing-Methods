Attribute VB_Name = "BoolTests"
Option Explicit
Option Private Module

'@TestModule
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed, LineLabelNotUsed
'@Folder("API.Utils.Tests")

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
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub DefaultPropertyLetGet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim b As New Bool
    b.Value = False
    
    'Act:
    b = True
    
    'Assert:
    Assert.AreEqual True, b.Value
    Assert.AreEqual True, (b) 'should just be b - this is an issue with the assert class

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub ClassConstructor()
    On Error GoTo TestFail
    
    'Arrange:
    Dim a As Bool, b As Bool, c As Bool
    
    'Act:
    Set b = Bool.Create(True)
    Set a = Bool.Create(False)
    Set c = Bool.Create(a)                       'implicit conversion with CBool
    
    'Assert:
    Assert.AreEqual True, b.Value
    Assert.AreEqual False, a.Value
    Assert.AreEqual a.Value, c.Value
    Assert.AreNotSame a, c                       'c only has the same value as a

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub AssigningByReferenceCanOverwrite()
    On Error GoTo TestFail
    
    'Arrange:
    Dim base As Bool, copy As Bool

    'Act:
    Set base = Bool.Create(True)
    Set copy = Bool.FromPtr(objPtr(base))
    copy = False

    'Assert:
    Assert.AreEqual False, base.Value
    Assert.AreSame base, copy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub InvalidConstructionRaisesTypeMismatchError()
    Const ExpectedError As Long = 13             'type mismatch
    On Error GoTo TestFail
    
    'Arrange:
    Dim b As Bool

    'Act:
    Set b = Bool.Create("Not a boolean!")

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

