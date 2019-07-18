Attribute VB_Name = "MessageQueueStuff"
'@Folder("Tests")
Option Explicit

Private Declare Function GetQueueStatus Lib "user32" (ByVal flags As Long) As Long

Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (ByVal lpMsg As LongPtr, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Boolean

Private Const QS_TIMER As Long = &H10

Private Sub t()
    Dim result As Long
    result = GetQueueStatus(QS_TIMER)
End Sub
