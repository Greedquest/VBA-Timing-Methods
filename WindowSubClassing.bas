Attribute VB_Name = "WindowSubClassing"
'@Folder("WinAPI")
Option Explicit
Option Private Module

Private Const className As String = "Static"

Public Function tryGetMessageWindow(ByVal windowName As String, ByRef outHandle As LongPtr) As Boolean

    outHandle = FindWindow(className, windowName) 'better than storing in persistent dict as handles may change (apparently)

    If outHandle = 0 Then
        Const HWND_MESSAGE As Long = (-3&)
        outHandle = CreateWindowEx(0, className, windowName, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, 0)
    End If
    
    tryGetMessageWindow = outHandle <> 0
    
End Function

Private Function trySubclassWindow(ByVal WindowProc As LongPtr, ByVal windowHandle As LongPtr) As Boolean
    Static subClassIDs As Dictionary             'windowHandle:Dict[windowproc:id]
    If subClassIDs Is Nothing Then Set subClassIDs = Cache.loadObject("subClassIDs", New Dictionary)
        
    Dim instanceID As Long
    'Only let one instance of each windowProc per windowHandle
    If Not subClassIDs.Exists(windowHandle) Then subClassIDs.Add windowHandle, New Dictionary
    Dim procDict As Scripting.Dictionary
    Set procDict = subClassIDs(windowHandle)
    If procDict.Exists(WindowProc) Then
        instanceID = procDict(WindowProc)
    Else
        instanceID = procDict.Count
        procDict.item(instanceID) = WindowProc
    End If
    
    If SetWindowSubclass(windowHandle, WindowProc, instanceID) Then
        trySubclassWindow = True
    End If
    
End Function

Public Function tryHookMessageHandler(ByVal WindowProc As LongPtr, ByVal windowName As String, ByRef outHandle As LongPtr) As Boolean
    If Not tryGetMessageWindow(windowName, outHandle) Then
        Exit Function
    ElseIf Not trySubclassWindow(WindowProc, outHandle) Then
        Exit Function
    Else
        tryHookMessageHandler = True
    End If
End Function

'@Description("Destroy a message window (or any other Static window) by name. Returns True if successful or no matching window. Returns False and outHandle set to handle if unable to destroy")
Public Function tryDestroyMessageWindowByName(ByVal windowName As String, Optional ByRef outHandle As LongPtr) As Boolean
Attribute tryDestroyMessageWindowByName.VB_Description = "Destroy a message window (or any other Static window) by name. Returns True if successful or no matching window. Returns False and outHandle set to handle if unable to destroy"
    Dim successful As Boolean
    outHandle = WinAPI.FindWindow(className, windowName)
    If outHandle <> 0 Then
        successful = WinAPI.DestroyWindow(outHandle) <> 0
        'set to 0 if destroyed to mark handle invalid
        If successful Then outHandle = 0
    Else
        successful = True
    End If
    tryDestroyMessageWindowByName = successful
End Function

