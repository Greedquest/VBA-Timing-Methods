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

Private Sub t()
    Debug.Print "Posting... "; PostMessage(Application.hwnd, WM_TIMER, 0, 0)
    
    Dim result As Long
    result = GetQueueStatus(QS_TIMER)
    Debug.Print "Status: "; result
End Sub

Public Sub t2()
    Dim endTime As Single
    endTime = 2 + timer
    Dim outMsg As tagMSG
    Do Until PeekMessage(outMsg, Application.hwnd, 0, 0, PM_NOREMOVE) Or timer > endTime
        PostMessage Application.hwnd, WM_TIMER, 0, 0
    Loop
    Debug.Print outMsg.lParam
    
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
