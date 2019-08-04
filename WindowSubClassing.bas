Attribute VB_Name = "WindowSubClassing"
'@Folder("SubClassing")
Option Explicit
Option Private Module

Public Type handle
    value As LongPtr
End Type

Public Type hWnd
    handle As handle
End Type

'Public Enum WindowStyle
'    HWND_MESSAGE = (-3&)
'End Enum

Private Declare Function APiCreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
                         ByVal dwExStyle As Long, ByVal className As String, ByVal windowName As String, _
                         ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, _
                         ByVal nWidth As Integer, ByVal nHeight As Integer, _
                         ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, _
                         ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As hWnd

Private Declare Function ApiDestroyWindow Lib "user32" Alias _
                         "DestroyWindow" (ByRef windowHandle As hWnd) As Boolean

Private Declare Function ApiFindWindow Lib "user32" Alias "FindWindowA" ( _
                         ByVal lpClassName As String, _
                         ByVal lpWindowName As String) As hWnd
                         
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

Public Function getMessageWindow(ByVal windowName As String) As hWnd

    Const className As String = "Static"
    Dim result As hWnd
    result = ApiFindWindow(className, windowName)
    If result.handle.value = 0 Then              'not found
        Const HWND_MESSAGE As Long = (-3&)
        result = APiCreateWindowEx(0, className, windowName, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, 0)
    End If
    getMessageWindow = result
    
End Function

Public Sub t()
    Dim window As hWnd
    window = getMessageWindow("Barry")
    window = getMessageWindow("Barry")
    Debug.Print ApiDestroyWindow(window)
    window = getMessageWindow("Barry")
    Debug.Print ApiDestroyWindow(window)
End Sub


