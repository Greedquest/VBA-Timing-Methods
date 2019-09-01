Attribute VB_Name = "SubclassingError"
'@Folder("Old.Tests.Experiments")
'Option Explicit
'
'Private safeState As Boolean
'
'Private Function subclassProc(ByVal hWnd As LongPtr, ByVal uMsg As WindowsMessage, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
'
'    If safeState Then 'NoOp, but the window is dead anyway
'        Debug.Print "MSG #"; uMsg 'will this even print, or have we interrupted repainting the thread?
'        subclassProc = WinAPI.DefSubclassProc(hWnd, uMsg, wParam, lParam)
'    End If
'
'End Function
'
'Private Sub timerProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerID As LongPtr, ByVal tickCount As Long)
'    Debug.Print "MSG #"; message, tickCount
'End Sub
'
'Sub createWindow()
'    'get window and subclass it
'    safeState = True
'    Static messageWindow As ModelessMessageWindow 'so it hovers around in memory
'    Debug.Print "Creating window"
'    If Not ModelessMessageWindow.tryCreate(messageWindow, AddressOf subclassProc) Then
'        Debug.Print "Couldn't get/subclass window"
'        Exit Sub
'    End If
'End Sub
'
'Sub createTimer()
'    Debug.Print "Starting timer"
'    If WinAPI.SetTimer(Application.hWnd, 1, 500, AddressOf timerProc) = 0 Then
'        Debug.Print "Couldn't make timer"
'    End If
'End Sub
'
'Sub uncreateTimer()
'    Debug.Print "Killing: "; IIf(WinAPI.killTimer(Application.hWnd, 1) = 0, "failure", "success")
'End Sub
'
'Sub nukeEverything()
'    safeState = False
'    End
'End Sub
'
'Sub checkPointer()
'    Debug.Print "Address: "; VBA.CLngPtr(AddressOf subclassProc)
'    End
'End Sub
