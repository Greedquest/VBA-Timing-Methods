Attribute VB_Name = "WindowSubClassing"
'@Folder("SubClassing")
Option Explicit
Option Private Module
                         
Private Function SubclassHelloWorldProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    Debug.Print "Oi"; uMsg
    Select Case uMsg
        Case PM_MY_MESSAGE
            Debug.Print "Hello World"
            SubclassHelloWorldProc = True
        Case Else
            SubclassHelloWorldProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End Select
End Function

Private Function tryGetMessageWindow(ByVal windowName As String, ByVal reset As Boolean, ByRef outHandle As LongPtr) As Boolean

    Const className As String = "Static"
    outHandle = FindWindow(className, windowName) 'better than storing in persistent dict as handles may change (apparently)
    If outHandle <> 0 And reset Then              'was found and needs resetting - i.e. destroying
        If WinAPI.DestroyWindow(outHandle) Then
            outHandle = 0 'invalidate handle so a new window is generated
        Else
            Exit Function 'couldn't destroy, let's ignore it...
        End If
    End If
    
    If outHandle = 0 Then
        Const HWND_MESSAGE As Long = (-3&)
        outHandle = CreateWindowEx(0, className, windowName, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, 0)
    End If
    
    tryGetMessageWindow = outHandle <> 0
    
End Function

Private Function trySubclassWindow(ByVal windowProc As LongPtr, ByVal windowHandle As LongPtr) As Boolean
    Static subClassIDs As Dictionary 'id:windowProc pairs
    If subClassIDs Is Nothing Then Set subClassIDs = Cache.loadObject("subClassIDs", New Dictionary)
    
    If SetWindowSubclass(windowHandle, windowProc, subClassIDs.Count) Then
        On Error Resume Next
        subClassIDs.Add subClassIDs.Count, windowProc 'NOTE never remove from this collection or id generation gets confused
        trySubclassWindow = Err.Number = 0
        On Error GoTo 0
    End If
    
End Function

Private Function tryHookMessageHandler(ByVal windowProc As LongPtr, ByVal windowName As String, ByRef outHandle As LongPtr) As Boolean
    If Not tryGetMessageWindow(windowName, outHandle) Then
        Exit Function
    ElseIf Not trySubclassWindow(windowProc, outHandle) Then
        Exit Function
    Else
        tryHookMessageHandler = True
    End If
End Function


