Attribute VB_Name = "WinAPI"
'@Folder("WinAPI")
Option Explicit
Option Private Module

Public Type tagPOINT
    X As Long
    Y As Long
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
                        ByVal flags As QueueStatusFlag) As Long

Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" ( _
                        ByRef lpMsg As tagMSG, _
                        ByVal hWnd As LongPtr, _
                        ByVal wMsgFilterMin As WindowsMessage, _
                        ByVal wMsgFilterMax As WindowsMessage, _
                        ByVal wRemoveMsg As PeekMessageFlag) As Boolean

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal msg As WindowsMessage, _
                        ByVal wParam As LongPtr, _
                        ByVal lParam As LongPtr) As Boolean

Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" ( _
                        ByRef lpMsg As tagMSG) As LongPtr

'Windows
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
                        ByVal dwExStyle As Long, ByVal className As String, ByVal windowName As String, _
                        ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, _
                        ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, _
                        ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr

Public Declare Function DestroyWindow Lib "user32" ( _
                        ByVal hWnd As LongPtr) As Boolean

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As LongPtr
                         
'Subclassing
Public Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal uMsg As WindowsMessage, _
                        ByVal wParam As LongPtr, _
                        ByVal lParam As LongPtr) As Boolean

Public Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal pfnSubclass As LongPtr, _
                        ByVal uIdSubclass As LongPtr, _
                        Optional ByVal dwRefData As LongPtr) As Boolean

Public Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal pfnSubclass As LongPtr, _
                        ByVal uIdSubclass As LongPtr) As Boolean

'Timers
Public Declare Function SetTimer Lib "user32" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal nIDEvent As Long, _
                        ByVal uElapse As Long, _
                        ByVal lpTimerFunc As LongPtr) As Long

Public Declare Function KillTimer Lib "user32" ( _
                        ByVal hWnd As LongPtr, _
                        ByVal nIDEvent As Long) As Long
                         
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                        ByVal lpPrevWndFunc As LongPtr, _
                        ByRef timerFlag As Bool, _
                        Optional ByVal message As WindowsMessage = WM_NOTIFY, _
                        Optional ByVal timerID As Long = 0, _
                        Optional ByVal unused3 As Long) As Long
