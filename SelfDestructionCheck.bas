Attribute VB_Name = "SelfDestructionCheck"
'@Folder("Tests")
'@IgnoreModule
Option Explicit

Private Function SelfDestructMessageWindowProc(ByVal hWnd As LongPtr, ByVal uMsg As WindowsMessage, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    
    Debug.Print "Message #"; uMsg
    If uMsg = WM_TIMER Then
        Debug.Print "Destroying from message window:"; WinAPI.DestroyWindow(hWnd)
    Else
        SelfDestructMessageWindowProc = WinAPI.DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Private Sub DestroyerProc(ByVal hWnd As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    Debug.Print "Destroying from TimerProc: "; WinAPI.DestroyWindow(hWnd)
    Debug.Print "Killing Timer from TimerProc:"; WinAPI.KillTimer(hWnd, timerID)
End Sub

Sub selfDestructProc()
    Dim handle As LongPtr
    If Not tryHookMessageHandler(AddressOf SelfDestructMessageWindowProc, "Quillam", handle) Then
        Debug.Print "Err make handler"
        Exit Sub
    End If
    Dim result As Long
    result = WinAPI.PostMessage(handle, WM_TIMER, 0, AddressOf DestroyerProc) <> 0
    Debug.Print "Post message: "; result
    'validation timer
    WinAPI.SetTimer handle, 0, &HFFFFFFFF, AddressOf DestroyerProc
    
    'WindowSubClassing.tryDestroyMessageWindowByName "Quillam"
End Sub

Sub t()
    Dim a As Boolean, b As Long
    Debug.Print LenB(a), LenB(b)
End Sub

