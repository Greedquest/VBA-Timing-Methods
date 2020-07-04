Attribute VB_Name = "TickerApiUnitTests"
Option Explicit
Option Private Module

'@TestModule
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed, LineLabelNotUsed
'@Folder("Tests")
                         
Private tempIDs As Collection                    'holds ids of all timers so they can be killed manually

Private Assert As Rubberduck.PermissiveAssertClass
'@Ignore VariableNotUsed: RD auto
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
    Dim ID As Variant
    For Each ID In tempIDs
        WinAPI.KillTimer TickerAPI.messageWindowHandle, ID
    Next ID
    Set TickerAPI = New TickerAPI                'the authentic way of killing stuff is just to reset the API
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Debug.Print String(50, "-")
    Set TickerAPI = New TickerAPI
    UnmanagedTimerTestProcs.clearLog
End Sub

'@TestCleanup
'@Ignore EmptyMethod: RD auto
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

'TODO write this test, it's a good one
''@TestMethod("Uncategorized")
'Private Sub StartExistingTimerNoError()
'    'TODO is there a way to test whether WinAPI.SetTimer on ObjPtr(callbackwrapper) *before* TickerAPI.Start[Un|Managed]Timer is bad?
'End Sub

'@TestMethod("Uncategorized")
Private Sub KillNonExistentTimerRaisesDestroyTimerError()
    Const ExpectedError As Long = TimerError.DestroyTimerError
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim timerID As LongPtr
    timerID = TickerAPI.StartUnmanagedTimer(AddressOf QuietNoOpCallback, False, INFINITE_DELAY)
    
    Dim killSuccess As Boolean
    killSuccess = WinAPI.KillTimer(TickerAPI.messageWindowHandle, timerID) <> 0
    
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
    Dim ID As LongPtr
    ID = TickerAPI.StartUnmanagedTimer(AddressOf QuietNoOpCallback, False, INFINITE_DELAY)
    
    'Assert:
    Assert.AreNotEqual CLng(0), CLng(ID)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("UnmanagedTimerExperiments")
Private Sub UnmanagedTimerImmediateCall()
    On Error GoTo TestFail
    
    'Arrange:
    Dim someData As String
    someData = "blah"
    
    'Act:
    Dim timerID As LongPtr
    timerID = TickerAPI.StartUnmanagedTimer(AddressOf UnmanagedTimerTestProc, True, data:=someData)
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

'@TestMethod("UnmanagedTimerExperiments")
Private Sub UnmanagedTimerDelayedCall()
    On Error GoTo TestFail
    
    'Arrange:
    Dim someData As String
    someData = "blah"
    
    'Act:
    Dim timerID As LongPtr
    timerID = TickerAPI.StartUnmanagedTimer(AddressOf UnmanagedTimerTestProc, False, data:=someData)
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

