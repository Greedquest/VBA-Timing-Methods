Attribute VB_Name = "WindowSubClassing"
'@Folder("SubClassing")
Option Explicit
Option Private Module

Private Const className As String = "Static"
                         
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

Public Function tryGetMessageWindow(ByVal windowName As String, ByRef outHandle As LongPtr) As Boolean

    outHandle = FindWindow(className, windowName) 'better than storing in persistent dict as handles may change (apparently)

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

Public Function tryHookMessageHandler(ByVal windowProc As LongPtr, ByVal windowName As String, ByRef outHandle As LongPtr) As Boolean
    If Not tryGetMessageWindow(windowName, outHandle) Then
        Exit Function
    ElseIf Not trySubclassWindow(windowProc, outHandle) Then
        Exit Function
    Else
        tryHookMessageHandler = True
    End If
End Function

'@Description("Destroy a message window (or any other Static window) by name. Returns True if successful or no matching window. Returns False and outHandle set to handle if unable to destroy")
Public Function tryDestroyMessageWindowByName(ByVal windowName As String, Optional ByRef outHandle As LongPtr) As Boolean
    Dim successful As Boolean
    outHandle = WinAPI.FindWindow(className, windowName)
    If outHandle <> 0 Then
        successful = WinAPI.DestroyWindow(handle)
        'set to 0 if destroyed to mark handle invalid
        If successful Then outHandle = 0
    Else
        successful = True
    End If
    tryDestroyMessageWindowByName = successful
End Function
