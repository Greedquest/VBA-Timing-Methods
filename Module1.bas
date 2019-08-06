Attribute VB_Name = "Module1"
'@Folder("Tests")
'@IgnoreModule
Option Explicit

Private Function SelfDestructMessageWindowProc(ByVal hWnd As LongPtr, ByVal uMsg As WindowsMessage, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    
    Debug.Print "Message #"; uMsg
    If uMsg = WM_TIMER Then
        Debug.Print WinAPI.DestroyWindow(hWnd)
    Else
        MessageWindowProc = WinAPI.DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Private Sub DestroyerProc(ByVal hWnd As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    WinAPI.DestroyWindow hWnd
End Sub

Sub selfDestructProc()
    Dim handle As LongPtr
    If Not tryHookMessageHandler(AddressOf SelfDestructMessageWindowProc, "Quillam", handle) Then
        Debug.Print "Err make handler"
        Exit Sub
    End If
    If Not WinAPI.PostMessage(handle, WM_TIMER, 0, AddressOf DestroyerProc) Then
        Debug.Print "Err post message"
    End If
    
    WindowSubClassing.tryDestroyMessageWindowByName "Quillam"
End Sub
