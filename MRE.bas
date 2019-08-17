Attribute VB_Name = "MRE"
Option Explicit

Private Sub timerProc(ByVal windowHandle As LongPtr, ByVal message As WindowsMessage, ByVal timerObj As Collection, ByVal tickCount As Long)
    Debug.Print timerObj.item("myItem")
    Static i As Long
    i = i + 1
    If i > 10 Then WinAPI.KillTimer Application.hWnd, ObjPtr(timerObj)
End Sub

Private Sub setUpTimer()
    Dim testCollection As Collection
    Set testCollection = Cache.loadObject("MRE.TestCollection", New Collection)
    testCollection.Add "Hi there!", "myItem"
    WinAPI.SetTimer Application.hWnd, ObjPtr(testCollection), 500, AddressOf timerProc
End Sub
