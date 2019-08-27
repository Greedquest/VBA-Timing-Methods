Attribute VB_Name = "WinAPI"
'@Folder("WinAPI")
Option Explicit
Option Private Module

Public Type tagPOINT
    X As Long
    Y As Long
End Type

Public Type DWORD 'same size as Long, but intellisense on members is nice
    LoWord As Integer
    HiWord As Integer
End Type

Public Type tagMSG
    hWnd As LongPtr
    message As WindowsMessage
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    cursor As tagPOINT
    #If Mac Then
    lPrivate As Long
    #End If
End Type

Public Type timerMessage
    windowHandle As LongPtr
    messageEnum As WindowsMessage
    timerID As LongPtr
    timerProc As LongPtr
    tickCountTime As Long
    cursor As tagPOINT
    #If Mac Then
    lPrivate As Long
    #End If
End Type

Public Type WNDCLASSEX
    cbSize         As Long
    style          As Long                       ' See CS_* constants
    lpfnwndproc    As LongPtr
    '   lpfnwndproc    As Long
    cbClsextra     As Long
    cbWndExtra     As Long
    hInstance      As LongPtr
    hIcon          As LongPtr
    hCursor        As LongPtr
    hbrBackground  As LongPtr
    '   hInstance      as long
    '   hIcon          as long
    '   hCursor        as long
    '   hbrBackground  as long
    lpszMenuName   As String
    lpszClassName  As String
    hIconSm        As LongPtr
    '   hIconSm        as long
End Type

Public Enum WindowStyle
    HWND_MESSAGE = (-3&)
End Enum

Public Enum QueueStatusFlag
    QS_TIMER = &H10
    QS_ALLINPUT = &H4FF
End Enum

Public Enum PeekMessageFlag
    PM_REMOVE = &H1
    PM_NOREMOVE = &H0
End Enum

''@Description("Windows Timer Message https://docs.microsoft.com/windows/desktop/winmsg/wm-timer")
Public Enum WindowsMessage
    WM_TIMER = &H113
    WM_NOTIFY = &H4E                             'arbitrary, sounds nice though
End Enum

'Messages
Public Declare Function GetQueueStatus Lib "user32" ( _
                        ByVal flags As QueueStatusFlag) As DWORD

Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" ( _
                        ByRef lpMsg As tagMSG, _
                        ByVal hWnd As LongPtr, _
                        ByVal wMsgFilterMin As WindowsMessage, _
                        ByVal wMsgFilterMax As WindowsMessage, _
                        ByVal wRemoveMsg As PeekMessageFlag) As Long
                        
Public Declare Function PeekTimerMessage Lib "user32" Alias "PeekMessageA" ( _
                        ByRef outMessage As timerMessage, _
                        ByVal hWnd As LongPtr, _
                        Optional ByVal wMsgFilterMin As WindowsMessage = WM_TIMER, _
                        Optional ByVal wMsgFilterMax As WindowsMessage = WM_TIMER, _
                        Optional ByVal wRemoveMsg As PeekMessageFlag = PM_REMOVE) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal msg As WindowsMessage, _
                        ByVal wParam As LongPtr, _
                        ByVal lParam As LongPtr) As Long

Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" ( _
                        ByVal lpMsg As LongPtr) As LongPtr
                        
Public Declare Function DispatchTimerMessage Lib "user32" Alias "DispatchMessageA" ( _
                        ByRef message As timerMessage) As LongPtr

'Windows
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
                        ByVal dwExStyle As Long, ByVal className As String, ByVal windowName As String, _
                        ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, _
                        ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, _
                        ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr

Public Declare Function DestroyWindow Lib "user32" ( _
                        ByVal hWnd As LongPtr) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As LongPtr
                         
'Registering

Public Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" ( _
                        ByRef pcWndClassEx As WNDCLASSEX) As Long
                        
                        
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" ( _
                        ByVal lpClassName As String, ByVal hInstance As LongPtr) As Long
                        
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" ( _
                        ByVal lhwnd As LongPtr, _
                        ByVal wMsg As Long, _
                        ByVal wParam As LongPtr, _
                        ByVal lParam As LongPtr) As Long


Public Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal uMsg As WindowsMessage, _
                        ByVal wParam As LongPtr, _
                        ByVal lParam As LongPtr) As Long

Public Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal pfnSubclass As LongPtr, _
                        ByVal uIdSubclass As LongPtr, _
                        Optional ByVal dwRefData As LongPtr) As Long

Public Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal pfnSubclass As LongPtr, _
                        ByVal uIdSubclass As LongPtr) As Long

'Timers
Public Declare Function SetTimer Lib "user32" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal nIDEvent As LongPtr, _
                        ByVal uElapse As Long, _
                        ByVal lpTimerFunc As LongPtr) As LongPtr

Public Declare Function killTimer Lib "user32" Alias _
                        "KillTimer" (ByVal hWnd As _
                        LongPtr, ByVal nIDEvent As LongPtr) As Long
                         
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                        ByVal lpPrevWndFunc As LongPtr, _
                        ByRef params As UnmanagedCallbackWrapper, _
                        Optional ByVal message As WindowsMessage = WM_NOTIFY, _
                        Optional ByVal timerID As Long = 0, _
                        Optional ByVal unused3 As Long) As LongPtr

Public Sub PrintMessageQueue(ByVal windowHandle As LongPtr, Optional ByVal filterLow As WindowsMessage = 0, Optional ByVal filterHigh As WindowsMessage = 0)
    Dim msg As tagMSG
    Dim results As New Dictionary
    Do While PeekMessage(msg, windowHandle, filterLow, filterHigh, PM_REMOVE) <> 0
        If results.Exists(msg.message) Then
            results(msg.message) = results(msg.message) + 1
        Else
            results(msg.message) = 1
        End If
    Loop
    'put them back?
    If results.Count = 0 Then
        Debug.Print "No Messages"
    Else
        Dim key As Variant
        For Each key In results.Keys
            Debug.Print "#"; key; ":", results(key)
        Next key
    End If
End Sub


