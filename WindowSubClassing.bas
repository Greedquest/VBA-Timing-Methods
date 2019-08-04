Attribute VB_Name = "WindowSubClassing"
'@Folder("SubClassing")
Option Explicit
Option Private Module


Private Declare Function APiCreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
                         ByVal dwExStyle As Long, ByVal className As String, ByVal windowName As String, _
                         ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, _
                         ByVal nWidth As Integer, ByVal nHeight As Integer, _
                         ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, _
                         ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr

Private Declare Function ApiDestroyWindow Lib "user32" Alias _
                         "DestroyWindow" (ByVal windowHandle As LongPtr) As Boolean

Private Declare Function ApiFindWindow Lib "user32" Alias "FindWindowA" ( _
                         ByVal lpClassName As String, _
                         ByVal lpWindowName As String) As LongPtr
                         
Private Declare Function ApiDefSubclassProc Lib "comctl32.dll" Alias "#413" ( _
                         ByVal hWnd As LongPtr, _
                         ByVal uMsg As Long, _
                         ByVal wParam As LongPtr, _
                         ByVal lParam As LongPtr) As Boolean

Private Declare Function ApiSetWindowSubclass Lib "comctl32.dll" Alias "#410" ( _
                         ByVal hWnd As LongPtr, _
                         ByVal pfnSubclass As LongPtr, _
                         ByVal uIdSubclass As LongPtr, _
                         Optional ByVal dwRefData As LongPtr) As Boolean

Private Declare Function ApiRemoveWindowSubclass Lib "comctl32.dll" Alias "#412" ( _
                         ByVal hWnd As LongPtr, _
                         ByVal pfnSubclass As LongPtr, _
                         ByVal uIdSubclass As LongPtr) As Boolean
                         
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

    Const className As String = "Static"
    outHandle = ApiFindWindow(className, windowName) 'better than storing in persistent dict as handles may change (apparently)
    If outHandle = 0 Then              'not found
        Const HWND_MESSAGE As Long = (-3&)
        outHandle = APiCreateWindowEx(0, className, windowName, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, 0)
    End If
    
    tryGetMessageWindow = outHandle <> 0
    
End Function


Public Sub trySubclassWindow(ByVal windowProc As LongPtr, ByVal windowHandle As LongPtr)
'    Static subClassIDs As Scripting.Dictionary
'    If subClassIDs Is Nothing Then Set subClassIDs = PersistentDict.Create("TimerAPI.Windows").Data
    
    If ApiSetWindowSubclass(windowHandle, windowProc, subClassIDs.Count) Then
        trySubclassWindow = True
'        subClassIDs.Add windowProc, subClassIDs.Count 'NOTE never remove or this gets confused
    End If
    
End Sub

Public Function tryHookMessageHandler(ByVal windowProc As LongPtr, ByVal windowName As String, ByRef outHandle As LongPtr) As Boolean
    If Not tryGetMessageWindow(windowName, outHandle) Then
        Exit Function
    ElseIf Not trySubclassWindow(windowProc, outHandle) Then
        Exit Function
    Else
        tryHookMessageHandler = True
    End If
End Function

Sub t()
    Dim handle As LongPtr
    If tryHookMessageHandler(AddressOf SubclassHelloWorldProc, "Tests", handle) Then
        PostMessage handle, WM_TIMER, 0, 0
        Debug.Print ApiDestroyWindow(handle)
    End If
End Sub
