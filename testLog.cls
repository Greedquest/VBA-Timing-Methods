VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "testLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests")
Option Explicit

Private Type testLogData
    callLog As New Scripting.Dictionary '{callerID:[data1,data2...]}
    errorLog As New Scripting.Dictionary '{callerID:[(errNum1,errInfo1), (errNum2,errInfo2),...]}
End Type

Private this As testLogData

Public Event afterLog()
Public Event afterError()

Public Sub logError(ByVal callerID As LongPtr, ByVal errNum As Long, ByVal errDescription As String, Optional ByVal Source As String = "TimerProc")
    If Not this.errorLog.Exists(callerID) Then
        this.errorLog.Add callerID, New Collection
    End If
    Dim errorSet As Collection
    Set errorSet = this.errorLog.Item(callerID)
    errorSet.Add Array(errNum, Toolbox.Strings.Format("{0} caused an error in {3}: #{1} - {2}", callerID, errNum, errDescription, Source))
End Sub

Public Sub logCall(ByVal callerID As LongPtr, ByVal data As Variant)
    If Not this.callLog.Exists(callerID) Then
        this.callLog.Add callerID, New Collection
    End If
    Dim callSet As Collection
    Set callSet = this.callLog.Item(callerID)
    callSet.Add data
End Sub

Public Property Get errorLog() As Dictionary
    Set errorLog = this.errorLog
End Property

Public Property Get callLog() As Dictionary
    Set callLog = this.callLog
End Property

Public Property Get errorCount(Optional ByVal callerID As LongPtr = 0) As Long
    Dim errCount As Long
    If callerID = 0 Then 'total
        Dim errorSet As Variant
        For Each errorSet In this.errorLog.Items
            errCount = errCount + errorSet.Count
        Next errorSet
    Else
        On Error Resume Next
        errCount = this.errorLog(callerID).Count
        On Error GoTo 0
    End If
    errorCount = errCount
End Property

Public Property Get callCount(Optional ByVal callerID As LongPtr = 0) As Long
    If callerID = 0 Then 'total
        Dim runningTot As Long
        Dim callSet As Variant
        For Each callSet In this.callLog.Items
            runningTot = runningTot + callSet.Count
        Next callSet
        callCount = runningTot
    Else
        On Error Resume Next
        callCount = this.callLog(callerID).Count
        On Error GoTo 0
    End If
End Property
