Attribute VB_Name = "WindowCreation"
'@Folder("WinAPI")
Option Explicit

Const className As String = "TimerAPIMessageWindow"
Const windowName As String = "Barry"

Private Function HelloWorldWindowProc(ByVal hWnd As LongPtr, ByVal uMsg As WindowsMessage, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Debug.Print "Message #"; uMsg
    If uMsg = PM_MY_MESSAGE Then
        Debug.Print "Hello World"
    End If
    HelloWorldWindowProc = WinAPI.DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
Sub registerClass()
    'static const char* class_name = "DUMMY_CLASS";
    
    Dim classStruct As WNDCLASSEX
    classStruct.cbSize = LenB(classStruct)
    classStruct.lpfnwndproc = VBA.CLngPtr(AddressOf HelloWorldWindowProc)
    classStruct.hInstance = Application.HinstancePtr
    classStruct.lpszClassName = className
    
    Debug.Print "Registering: "; className, WinAPI.RegisterClassEx(classStruct) <> 0

End Sub

Function makeWindow() As LongPtr
    Dim hWnd As LongPtr
    hWnd = WinAPI.FindWindow(className, windowName)
    
    If hWnd = 0 Then
        Debug.Print GetSystemErrorMessageText(Err.LastDllError)
        hWnd = WinAPI.CreateWindowEx(0, className, windowName, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, Application.HinstancePtr, 0)
    End If
    If hWnd = 0 Then
        Debug.Print "Errorroororororor "; GetSystemErrorMessageText(Err.LastDllError)
    Else
        Debug.Print IIf(PostMessage(hWnd, PM_MY_MESSAGE, 0, 0) <> 0, "Message sent", "Couldn't send:( " & GetSystemErrorMessageText(Err.LastDllError))
    End If
    makeWindow = hWnd
End Function

Sub clicktimer()
    Dim hWnd As LongPtr
    hWnd = makeWindow
    Debug.Print hWnd, IIf(SetTimer(hWnd, 11, 1000, AddressOf CallbackFunctions.RawSelfKillingProc) <> 0, "made a timer", "Couldn't send:( " & GetSystemErrorMessageText(Err.LastDllError))
End Sub

