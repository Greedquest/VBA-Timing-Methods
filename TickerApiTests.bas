Attribute VB_Name = "TickerApiTests"
Option Explicit
Option Private Module

'@TestModule
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed, LineLabelNotUsed
'@Folder("Tests")

Private Declare Function ApiSetTimer Lib "user32" Alias "SetTimer" ( _
                         ByVal hWnd As Long, _
                         ByVal nIDEvent As Long, _
                         ByVal uElapse As Long, _
                         ByVal lpTimerFunc As Long) As Long

Private Declare Function ApiKillTimer Lib "user32" Alias "KillTimer" ( _
                         ByVal hWnd As Long, _
                         ByVal nIDEvent As Long) As Long
                         
Private tempIDs As Dictionary

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
    Set tempIDs = New Dictionary
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Dim id As Variant
    For Each id In tempIDs.Keys
        ApiKillTimer Application.hWnd, id
    Next id
    TickerAPI.killAllTimers
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Debug.Print String(50, "-")
    Set TickerAPI = New TickerAPI
    TickerAPI.UnlockApi                          'for simplicity
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

'@TestMethod("Uncategorized")
Private Sub StartExistingTimerNoError()
    On Error GoTo TestFail
    
    'Arrange:
    tempIDs.Add 1, ApiSetTimer(Application.hWnd, 1, 10000, AddressOf SafeCallbackProc)
            
    'Act:
    Dim apiID As Long
    apiID = TickerAPI.StartTimer(AddressOf SafeCallbackProc, False)

    'Assert:
    Assert.AreEqual 1&, apiID, printf("Expected {0}, actual {1}", 1, apiID)


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub KillNonExistentTimerRaisesDestroyTimerError()
    Const ExpectedError As Long = DestroyTimerError
    On Error GoTo TestFail
    
    'Arrange:
    TickerAPI.StartTimer AddressOf QuietNoOpCallback, False
    Dim killSuccess As Boolean
    killSuccess = ApiKillTimer(TickerAPI.messageWindowHandle, TickerAPI.StartTimer(AddressOf QuietNoOpCallback, False))
    
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


