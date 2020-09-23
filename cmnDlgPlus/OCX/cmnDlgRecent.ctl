VERSION 5.00
Begin VB.UserControl cmnDlgRecent 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   Picture         =   "cmnDlgRecent.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   465
   ToolboxBitmap   =   "cmnDlgRecent.ctx":08CA
End
Attribute VB_Name = "cmnDlgRecent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'Nothing special here - just talking to the module
Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Width = 480
    UserControl.Height = 480
End Sub

Public Property Get Filter() As String
    Filter = cmnDlg.Filter
End Property

Public Property Let Filter(ByVal vNewValue As String)
    cmnDlg.Filter = vNewValue
End Property

Public Property Get Filterindex() As Long
    Filterindex = cmnDlg.Filterindex
End Property

Public Property Let Filterindex(ByVal vNewValue As Long)
    cmnDlg.Filterindex = vNewValue
End Property

Public Property Get FileTitle() As String
    FileTitle = cmnDlg.FileTitle
End Property

Public Property Let FileTitle(ByVal vNewValue As String)
    cmnDlg.FileTitle = vNewValue
End Property

Public Property Get FileName() As String
    FileName = cmnDlg.FileName
End Property

Public Property Let FileName(ByVal vNewValue As String)
    cmnDlg.FileName = vNewValue
End Property

Public Property Get InitDir() As String
    InitDir = cmnDlg.InitDir
End Property

Public Property Let InitDir(ByVal vNewValue As String)
    cmnDlg.InitDir = vNewValue
End Property

Public Property Get DialogTitle() As String
    DialogTitle = cmnDlg.DialogTitle
End Property

Public Property Let DialogTitle(ByVal vNewValue As String)
    cmnDlg.DialogTitle = vNewValue
End Property

Public Property Get OwnerFrm() As Object
    Set OwnerFrm = cmnDlg.OwnerFrm
End Property

Public Property Let OwnerFrm(ByVal vNewValue As Object)
    Set cmnDlg.OwnerFrm = vNewValue
End Property

Public Property Get Flags() As Long
    Flags = cmnDlg.Flags
End Property

Public Property Let Flags(ByVal vNewValue As Long)
    cmnDlg.Flags = vNewValue
End Property

Public Property Get Dorecent() As Boolean
    Dorecent = cmnDlg.Dorecent
End Property

Public Property Let Dorecent(ByVal vNewValue As Boolean)
    cmnDlg.Dorecent = vNewValue
End Property

Public Property Get OpenRecent() As Boolean
    OpenRecent = cmnDlg.OpenRecent
End Property

Public Property Let OpenRecent(ByVal vNewValue As Boolean)
    cmnDlg.OpenRecent = vNewValue
End Property

Public Property Get AddTorecentDocs() As Boolean
    AddTorecentDocs = cmnDlg.AddTorecentDocs
End Property

Public Property Let AddTorecentDocs(ByVal vNewValue As Boolean)
    cmnDlg.AddTorecentDocs = vNewValue
End Property

Public Sub AddToMRUList(mfile As String)
    AddMRU mfile
End Sub

Public Sub RemoveFromMRUList(mfile As String)
    RemoveMRU mfile
End Sub

Public Sub ClearMRUList()
    ClearMRUs
End Sub

Public Sub ShowOpen()
    UserControl.Parent.Enabled = False
    OwnerHandle = UserControl.hwnd
    OpenFile
    UserControl.Parent.Enabled = True
End Sub
Public Sub ShowSave()
    OwnerHandle = UserControl.hwnd
    SaveFile
End Sub

