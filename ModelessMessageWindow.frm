VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModelessMessageWindow 
   Caption         =   "UserForm1"
   ClientHeight    =   3072
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   4608
   OleObjectBlob   =   "ModelessMessageWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModelessMessageWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Timing")
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByRef outHwnd As LongPtr) As Long
#Else
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByRef outHwnd As Long) As Long
#End If

#If VBA7 Then
    Public Property Get hWnd() As LongPtr
        IUnknown_GetWindow Me, hWnd
    End Property
#Else
    Public Property Get hWnd() As Long
        IUnknown_GetWindow Me, hWnd
    End Property
#End If
