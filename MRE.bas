Attribute VB_Name = "MRE"
Option Explicit

Private Declare Function SetTimer Lib "user32" ( _
                         ByVal hWnd As LongPtr, _
                         ByVal nIDEvent As LongPtr, _
                         ByVal uElapse As Long, _
                         ByVal lpTimerFunc As LongPtr) As LongPtr

Private Declare Function KillTimer Lib "user32" ( _
                         ByVal hWnd As LongPtr, _
                         ByVal nIDEvent As LongPtr) As Long
                         
Private Function GetPersistentDictionary() As Object
    ' References:
    '  mscorlib.dll
    '  Common Language Runtime Execution Engine

    Const name = "weak-data"
    Static dict As Object

    If dict Is Nothing Then
        Dim host As New mscoree.CorRuntimeHost
        Dim domain As mscorlib.AppDomain
        host.Start
        host.GetDefaultDomain domain

        If IsObject(domain.GetData(name)) Then
            Set dict = domain.GetData(name)
        Else
            Set dict = CreateObject("Scripting.Dictionary")
            domain.SetData name, dict
        End If
    End If

    Set GetPersistentDictionary = dict
End Function

Private Sub timerProc(ByVal windowHandle As LongPtr, ByVal message As Long, ByVal timerObj As Object, ByVal tickCount As Long)
    Static i As Long
    i = i + 1
    Debug.Print i;
    Dim data As String
    Dim ptr As LongPtr
    ptr = ObjPtr(timerObj)
    On Error GoTo cleanFail
    data = timerObj.Item("myVal")
    On Error GoTo 0
    Debug.Print data
    If i >= 10 Then
cleanExit:
        'KillTimer Application.hWnd, ptr ' ObjPtr(timerObj)
        Debug.Print "Done"

        i = 0
    End If
    Exit Sub
    
cleanFail:
    
    Set timerObj = Nothing
    Set GetPersistentDictionary().Item("testObj") = Nothing
    Resume cleanExit
End Sub

Private Sub setUpTimer()
    Dim cache As Dictionary
    Set cache = GetPersistentDictionary()
    Dim testObj As Object
    Set testObj = New fakeDictionary
    testObj.Item("myVal") = "I'm the data you passed!"
    Set cache.Item("testObj") = testObj
    SetTimer Application.hWnd, ObjPtr(testObj), 500, AddressOf timerProc
End Sub
