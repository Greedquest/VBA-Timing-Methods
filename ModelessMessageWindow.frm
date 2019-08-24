VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModelessMessageWindow 
   Caption         =   "UserForm1"
   ClientHeight    =   3072
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   4608
   OleObjectBlob   =   "ModelessMessageWindow.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModelessMessageWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("FirstLevelAPI")
Option Explicit

Private Type messageWindowData
    subClassIDs As New Dictionary '{proc:id}
End Type
Private this As messageWindowData

#If VBA7 Then
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByRef outHwnd As LongPtr) As Long
#Else
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByRef outHwnd As Long) As Long
#End If

#If VBA7 Then
    Public Property Get handle() As LongPtr
        IUnknown_GetWindow Me, handle
    End Property
#Else
    Public Property Get handle() As Long
        IUnknown_GetWindow Me, handle
    End Property
#End If

Public Function tryCreate(ByRef outWindow As ModelessMessageWindow, Optional ByVal windowProc As LongPtr = 0, Optional ByVal data As LongPtr) As Boolean
    With New ModelessMessageWindow
        .Init
        If windowProc = 0 Then
            tryCreate = True
        Else
            tryCreate = .tryAddSubclass(windowProc, data)
        End If
'        .Hide
'        Debug.Print "------------------"
'        '.Hide
        Set outWindow = .Self
    End With
End Function

Public Property Get Self() As ModelessMessageWindow
    Set Self = Me
End Property

Public Sub Init()
    'Need to run this for window to be able to receive messages
    'Me.Show
    'Me.Hide
End Sub

Public Function tryAddSubclass(ByVal subclassProc As LongPtr, Optional ByVal data As LongPtr) As Boolean
        
    Dim instanceID As Long
    'Only let one instance of each subclassProc per windowHandle

    If this.subClassIDs.Exists(subclassProc) Then
        instanceID = this.subClassIDs(subclassProc)
    Else
        instanceID = this.subClassIDs.Count
        this.subClassIDs(subclassProc) = instanceID
    End If
    
    If WinAPI.SetWindowSubclass(handle, subclassProc, instanceID, data) Then
        tryAddSubclass = True
    End If
End Function

'@Description("Remove any registered subclasses - returns True if all removed successfully")
Public Function tryRemoveAllSubclasses() As Boolean
    
    Dim timerProc As Variant
    Dim result As Boolean
    result = True 'if no subclasses exist the we removed them nicely
    For Each timerProc In this.subClassIDs.Keys
        result = result And WinAPI.RemoveWindowSubclass(handle, timerProc, this.subClassIDs(timerProc)) <> 0
    Next timerProc
    this.subClassIDs.RemoveAll
    tryRemoveAllSubclasses = result
End Function

