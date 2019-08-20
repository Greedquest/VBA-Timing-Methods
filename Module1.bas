Attribute VB_Name = "Module1"
Option Explicit

Sub testcasting()

Debug.Print "Initialize"
Dim a As Object
Dim b As ICallbackWrapper
Dim c As Object
Debug.Print Toolbox.Strings.Format("\ta: {0} - {1}\n\tb: {2} - {3}\n\tc: {4} - {5}", ObjPtr(a), TypeName(a), ObjPtr(b), TypeName(b), ObjPtr(c), TypeName(c))

Debug.Print "DownCast"
Set a = New UnmanagedCallbackWrapper
Set b = a
Debug.Print Toolbox.Strings.Format("\ta: {0} - {1}\n\tb: {2} - {3}\n\tc: {4} - {5}", ObjPtr(a), TypeName(a), ObjPtr(b), TypeName(b), ObjPtr(c), TypeName(c))

Debug.Print "Upcast"
Set c = b
Debug.Print Toolbox.Strings.Format("\ta: {0} - {1}\n\tb: {2} - {3}\n\tc: {4} - {5}", ObjPtr(a), TypeName(a), ObjPtr(b), TypeName(b), ObjPtr(c), TypeName(c))

End Sub

Private Sub RawSafeTickingProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
    Static i As Long
    i = i + 1
    Debug.Print i; "Tick"
    
    If i >= 10 Then
        Debug.Print "Terminating"
        WinAPI.KillTimer windowHandle, timerID
    End If
End Sub

Private Function testMessageWindowSubclassProc(ByVal hWnd As LongPtr, ByVal uMsg As WindowsMessage, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    
    Debug.Print "Message #"; uMsg
    Select Case uMsg
    
            'NOTE this will never receive timer messages where TIMERPROC is specified,
        Case WindowsMessage.WM_TIMER             'wParam = timerID , lParam = "timerProc" (will be Null if it reaches here)
            Static i As Long
            i = i + 1
            If i >= 10 Then WinAPI.KillTimer hWnd, wParam
            Debug.Print i; "Loop tick"; dwRefData.Caption
            testMessageWindowSubclassProc = True
        Case Else
            testMessageWindowSubclassProc = WinAPI.DefSubclassProc(hWnd, uMsg, wParam, lParam)
            
    End Select
End Function

Sub makeTimer()
    Dim messageWindow As New ModelessMessageWindow
    messageWindow.Init
    If messageWindow.tryAddSubclass(AddressOf testMessageWindowSubclassProc, ObjPtr(messageWindow)) Then
        Debug.Print WinAPI.SetTimer(messageWindow.windowHandle, ObjPtr(messageWindow), 500, 0)
        'Debug.Print WinAPI.SetTimer(messageWindow., ObjPtr(messageWindow), 500, AddressOf RawSafeTickingProc)
    End If
End Sub
