Attribute VB_Name = "MessageQueueStuff"
'@Folder("Tests")
Option Explicit

Public Const PM_REMOVE As Long = &H1
Public Const PM_NOREMOVE As Long = &H0

Public Type tagPOINT
    x As Long
    y As Long
End Type

Public Type tagMSG
    hwnd As LongPtr
    message As Long
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    cursor As tagPOINT
    #If Mac Then
    lPrivate As Long
    #End If
End Type

Public Declare Function GetQueueStatus Lib "user32" (ByVal flags As Long) As Long

Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (ByRef lpMsg As tagMSG, ByVal hwnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Boolean
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Boolean

Public Const QS_TIMER As Long = &H10
Public Const QS_ALLINPUT As Long = &H4FF
Private Const WM_TIMER As Long = &H113

Function tryScheduleProc(timerProc As LongPtr, ByVal arg As Object) As Boolean
    Debug.Print "Scheduling..."
    tryScheduleProc = PostMessage(Application.hwnd, WM_TIMER, objPtr(arg), timerProc)
End Function

Sub asyncProc(ByVal hwnd As Long, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    Debug.Print "asyncProc called (should be called second)"
End Sub

Sub syncProc()
    Debug.Print "syncProc called (should be called first)"
End Sub

Sub test()
    If tryScheduleProc(AddressOf asyncProc, New Collection) Then
        syncProc
    Else
        Debug.Print "Error trying to schedule proc"
    End If
End Sub

Public Sub PrintMessageQueue(Optional filterLow As Long = 0, Optional filterHigh As Long = 0)
    Dim msg As tagMSG
    Dim results As New Dictionary
    Do While PeekMessage(msg, Application.hwnd, filterLow, filterHigh, PM_REMOVE)
        If results.Exists(msg.message) Then
            results(msg.message) = results(msg.message) + 1
        Else
            results(msg.message) = 1
        End If
    Loop
    'put them back?
    If results.Count = 0 Then
        Debug.Print "No Messages"
    Else
        Dim key
        For Each key In results.Keys
            Debug.Print key, results(key)
        Next key
    End If
End Sub
