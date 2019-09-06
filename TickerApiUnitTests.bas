Attribute VB_Name = "TickerApiUnitTests"
Option Explicit
Option Private Module

'@TestModule
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed, LineLabelNotUsed
'@Folder("Tests")
                         
Private tempIDs As Collection                    'holds ids of all timers so they can be killed manually
Private log As testLog

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
    Debug.Print String(50, "-")
    Set Assert = Nothing
    Set Fakes = Nothing
    Dim id As Variant
    For Each id In tempIDs
        WinAPI.killTimer TickerAPI.messageWindowHandle, id
    Next id
    Set TickerAPI = New TickerAPI                'the authentic way of killing stuff is just to reset the API
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Debug.Print String(50, "-")
    Set TickerAPI = New TickerAPI
    UnmanagedTimerTestProcs.clearLog
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
    Const ExpectedError As Long = TimerError.DestroyTimerError
    On Error GoTo TestFail
    
    'Arrange:
    'TODO infinite delay
    Dim timerID As LongPtr
    timerID = TickerAPI.StartUnmanagedTimer(AddressOf QuietNoOpCallback, False, INFINITE_DELAY)
    
    Dim killSuccess As Boolean
    killSuccess = WinAPI.killTimer(TickerAPI.messageWindowHandle, timerID) <> 0
    
    'Act:
    TickerAPI.KillTimerByID timerID

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

'@TestMethod("Uncategorized")
Private Sub StartUnmanagedTimerRaisesNoError()
    On Error GoTo TestFail
    

    'Act:
    Dim id As LongPtr
    id = TickerAPI.StartUnmanagedTimer(AddressOf QuietNoOpCallback, False, INFINITE_DELAY)
    
    'Assert:
    Assert.AreNotEqual 0, id

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("UnmanagedTimerExperiments")
Private Sub UnmanagedTimerImmediateCall()                'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    Dim someData As String
    someData = "blah"
    
    'Act:
    Dim timerID As LongPtr
    timerID = TickerAPI.StartUnmanagedTimer(AddressOf UnmanagedTimerTestProc, data:=someData)
    testLog.waitUntilTrigger
    TickerAPI.KillTimerByID timerID
    
    'Assert:
    Assert.AreEqual CLng(1), testLog.callCount(timerID), "Wrong number of calls"
    Assert.AreEqual CLng(0), testLog.errorCount(timerID), "Wrong number of errors"
    Assert.AreEqual someData, testLog.callLog(timerID)(1), "Data not right"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
