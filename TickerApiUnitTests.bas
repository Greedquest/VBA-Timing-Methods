Attribute VB_Name = "TickerApiUnitTests"
Option Explicit
Option Private Module

'@TestModule
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed, LineLabelNotUsed
'@Folder("Tests")
                         
Private tempIDs As Collection 'holds ids of all timers so they can be killed manually

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
    Set tempIDs = New Collection
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Dim id As Variant
    For Each id In tempIDs
        WinAPI.KillTimer TickerAPI.messageWindowHandle, id
    Next id
    Set TickerAPI = New TickerAPI 'the authentic way of killing stuff is just to reset the API
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Debug.Print String(50, "-")
    Set TickerAPI = New TickerAPI
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestKillByIdInvalidIdRaisesTimerNotFoundError()
    Const ExpectedError As Long = TimerError.TimerNotFoundError
    On Error GoTo TestFail

    'Arrange:

    'Act:
    TickerAPI.KillTimerByID 100

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
Private Sub KillByInvalidFunctionNoError()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    TickerAPI.KillTimersByFunction 100

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''@TestMethod("Uncategorized")
Private Sub StartExistingTimerNoError()
    'TODO is there a way to test whether WinAPI.SetTimer on ObjPtr(callbackwrapper) *before* TickerAPI.Start*Timer is bad?
End Sub

'@TestMethod("Uncategorized")
Private Sub KillNonExistentTimerRaisesDestroyTimerError()
    Const ExpectedError As Long = DestroyTimerError
    On Error GoTo TestFail
    
    'Arrange:
    TickerAPI.StartUnmanagedTimer AddressOf QuietNoOpCallback, False, 100000000 'TODO infinite delay in unit tests
    Dim killSuccess As Boolean
    killSuccess = WinAPI.KillTimer(TickerAPI.messageWindowHandle, TickerAPI.StartUnmanagedTimer(AddressOf QuietNoOpCallback, False)) <> 0
    
    'Act:
    TickerAPI.KillTimersByFunction AddressOf QuietNoOpCallback 'kill before it returns, but is already gone

Assert:
    Assert.IsTrue killSuccess, "Direct call did not kill the api"
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError And killSuccess Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


