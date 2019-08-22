VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Old.SubClassing")
'@IgnoreModule

Option Explicit
'@PredeclaredId

Private Const default_className As String = "Static"
Private Const default_windowName As String = "TickerAPI"
    
Private Declare Function APiCreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
                         ByVal dwExStyle As Long, _
                         ByVal lpClassName As String, _
                         ByVal lpWindowName As String, _
                         ByVal dwStyle As WindowStyle, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal nWidth As Long, _
                         ByVal nHeight As Long, _
                         ByVal hWndParent As LongPtr, _
                         ByVal hMenu As LongPtr, _
                         ByVal hInstance As LongPtr, _
                         ByVal lpParam As LongPtr) As LongPtr

Private Declare Function ApiDestroyWindow Lib "user32" Alias _
                         "DestroyWindow" (ByVal hWnd As Long) As Boolean

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
                         

Private Type tApiWindow
    className As String
    windowName As String
    handle As LongPtr
    style As WindowStyle
End Type

Private this As tApiWindow

Public Property Get style() As WindowStyle
    style = this.style
End Property

Public Property Let style(ByVal value As WindowStyle)
    this.style = value
End Property

Public Property Get handle() As LongPtr
    handle = this.handle
End Property

Public Property Get windowName() As String
    windowName = this.windowName
End Property

Public Property Let windowName(ByVal value As String)
    this.windowName = value
End Property

Public Property Get className() As String
    className = this.className
End Property

Public Property Let className(ByVal value As String)
    this.className = value
End Property

Public Function Self() As ApiWindow
    Set Self = Me
End Function

Public Function Create(Optional ByVal newClassName As String = default_className, Optional ByVal newWindowName As String = default_windowName, Optional ByVal newStyle As WindowStyle = HWND_MESSAGE) As ApiWindow
    With New ApiWindow
        .className = newClassName
        .style = newStyle
        .windowName = newWindowName
        If Not .tryGenerateWindow Then
            Exit Function
        End If
        Set Create = .Self
    End With
End Function

Public Function tryGenerateWindow() As Boolean
    this.handle = ApiFindWindow(this.className, this.windowName)
    If this.handle = 0 Then
        this.handle = APiCreateWindowEx(0, this.className, this.windowName, this.style, 0, 0, 0, 0, 0, 0, 0, 0)
    End If
    tryGenerateWindow = this.handle <> 0
End Function

Public Function tryAddSubclass(ByVal subclassProc As LongPtr, ByRef outSubclassID As LongPtr) As Boolean
    outSubclassID = getUniqueID
    tryAddSubclass = ApiSetWindowSubclass(this.handle, subclassProc, outSubclassID) <> 0
End Function

Public Function tryRemoveSubclass(ByVal subclassProc As LongPtr, ByVal subClassID As LongPtr) As Boolean
    tryRemoveSubclass = ApiRemoveWindowSubclass(this.handle, subclassProc, subClassID)
End Function

Private Function getUniqueID() As Long
    'TODO if we keep track of ids then they should always be consecutive with max 1 gap
    Static i As Long
    i = i + 1
    getUniqueID = i
End Function

Private Function tryDestroyWindow() As Boolean
    tryDestroyWindow = ApiDestroyWindow(this.handle)
End Function

Private Sub Class_Terminate()
    tryDestroyWindow
End Sub


