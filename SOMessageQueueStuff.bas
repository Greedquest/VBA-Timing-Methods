Attribute VB_Name = "SOMessageQueueStuff"
Option Explicit

'@Folder("Tests")

Private Sub asyncProc(ByVal hWnd As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    Debug.Print "asyncProc called (should be called second)"
    killTimer hWnd, timerID
End Sub

Private Sub syncProc()
    Debug.Print "syncProc called (should be called first)"
End Sub

Sub test()
    If tryScheduleProc(AddressOf asyncProc, New Collection) Then
        syncProc
    Else
        Debug.Print "Unable to schedule proc"
    End If
End Sub

Private Function tryScheduleProc(ByVal timerProc As LongPtr, ByVal arg As Object) As Boolean
    Debug.Print "Scheduling..."
    'make a validation timer - this won't expire for a long time
    'Debug.Print "Create a validation timer:"; SetTimer(Application.hWnd, objPtr(arg), &H7FFFFFFF, timerProc)
    tryScheduleProc = PostMessage(Application.hWnd, WM_TIMER, objPtr(arg), timerProc)
    PrintMessageQueue
End Function
