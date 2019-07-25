Attribute VB_Name = "SubClassing"
'@Folder("Tests.Subclassing")
Option Explicit

Private messageHwnd As LongPtr

Public Const PM_MY_MESSAGE As Long = &H400 + 1

Private Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Boolean
Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As Boolean
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Boolean

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare Function PostLongMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long

Private Function SubclassHelloWorldProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    Debug.Print "Oi"; uMsg
    Select Case uMsg
        Case PM_MY_MESSAGE
            Debug.Print "Hello World"
            SubclassHelloWorldProc = True
        Case Else
            SubclassHelloWorldProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
    End Select
End Function

Sub startSubclassing()
    If Not tryGetMessageWindow(messageHwnd) Then
        Debug.Print "Couldn't make a handle: "; messageHwnd
        Exit Sub
    End If
    If SetWindowSubclass(messageHwnd, AddressOf SubclassHelloWorldProc, 1) Then
        Debug.Print "Started fine"
    End If
    startTicking
End Sub


Sub stopSubclassing()
    'KillTimer messageHwnd, 1
    Dim hwnd As LongPtr
    hwnd = FindWindow("STATIC", "Barry")
    Debug.Print hwnd, messageHwnd
    If RemoveWindowSubclass(messageHwnd, AddressOf SubclassHelloWorldProc, 1) Then
        Debug.Print "Ended fine"
    End If
    Debug.Print DestroyWindow(messageHwnd)
End Sub

Sub sendMessage()
    Debug.Print messageHwnd, PM_MY_MESSAGE
    Debug.Print PostLongMessage(messageHwnd, PM_MY_MESSAGE, 0, 0)
End Sub

Sub startTicking()
    SetTimer messageHwnd, 1, 1000, 0
End Sub

Public Function tryGetMessageWindow(ByRef outHwnd As LongPtr) As Boolean
    Const HWND_MESSAGE As Long = (-3&)
    outHwnd = CreateWindowEx(0, "STATIC", "Barry", 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, 0)
    If outHwnd <> 0 Then tryGetMessageWindow = True
End Function
