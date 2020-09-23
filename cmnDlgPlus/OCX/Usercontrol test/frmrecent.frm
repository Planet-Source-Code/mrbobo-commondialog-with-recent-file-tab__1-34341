VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmrecent 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   6495
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "frmrecent.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCloseRight 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6240
      Picture         =   "frmrecent.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   45
      Width           =   210
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Existing"
      TabPicture(0)   =   "frmrecent.frx":070C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Recent"
      TabPicture(1)   =   "frmrecent.frx":0728
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LV"
      Tab(1).Control(1)=   "cmdOpen"
      Tab(1).Control(2)=   "cmdCancel"
      Tab(1).Control(3)=   "PicTempS"
      Tab(1).Control(4)=   "PicTempL"
      Tab(1).Control(5)=   "ImageListS"
      Tab(1).Control(6)=   "ImageListL"
      Tab(1).ControlCount=   7
      Begin MSComctlLib.ImageList ImageListL 
         Left            =   -71040
         Top             =   4440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   -2147483639
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageListS 
         Left            =   -71760
         Top             =   4440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   -2147483639
         _Version        =   393216
      End
      Begin VB.PictureBox PicTempL 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -72360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   4560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox PicTempS 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -72840
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   5
         Top             =   4560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70200
         TabIndex        =   4
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70200
         TabIndex        =   3
         Top             =   2760
         Width           =   1215
      End
      Begin MSComctlLib.ListView LV 
         Height          =   1935
         Left            =   -74640
         TabIndex        =   2
         Top             =   720
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Folder"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   240
         ScaleHeight     =   153
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   393
         TabIndex        =   1
         Top             =   480
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmrecent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To get the correct icons for the listview
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type
Enum SHGFI_FLAGS
    SHGFI_SMALLICON = &H1
    SHGFI_LARGEICON = &H0
    SHGFI_ICON = &H100
End Enum
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_FLAGS) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'flat listview column headers API
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const HDS_BUTTONS As Long = &H2
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE

Private Sub cmdCloseRight_Click()
    Unload Me 'equivalent to the "X" button
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    Dim temp As String 'return selection to cmnDlg type
    temp = LV.SelectedItem.SubItems(1)
    If Right(temp, 1) <> "\" Then temp = temp + "\"
    cmnDlg.OpenRecent = True
    cmnDlg.FileName = temp + LV.SelectedItem.Text
    cmnDlg.FileTitle = LV.SelectedItem.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Dim temp As String, z As Long
    Dim hHeader As Long
    'flat listview colomn headers thanks
    hHeader = SendMessage(LV.hwnd, LVM_GETHEADER, 0, ByVal 0&)
    SetWindowLong hHeader, GWL_STYLE, GetWindowLong(hHeader, GWL_STYLE) Xor HDS_BUTTONS
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DoDrag Me.hwnd 'allow dragging form
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Dim R As RECT
    LV.Height = SSTab1.Height * Screen.TwipsPerPixelY - 1880
    LV.Width = SSTab1.Width * Screen.TwipsPerPixelX - 840
    LV.ColumnHeaders(2).Width = LV.Width - LV.ColumnHeaders(1).Width - 120
    cmdOpen.Top = LV.Top + LV.Height + 105
    cmdOpen.Left = LV.Left + LV.Width - cmdOpen.Width
    cmdCancel.Left = LV.Left + LV.Width - cmdCancel.Width
    cmdCancel.Top = cmdOpen.Top + cmdOpen.Height + 60
    SetRect R, 1, 0, Me.ScaleWidth - 2, 18
    Me.Cls
    'paint titlebar
    DrawCaption Me.hwnd, Me.hDC, R, DC_ACTIVE Or DC_ICON Or DC_TEXT Or DC_GRADIENT
    cmdCloseRight.Left = Me.ScaleWidth - 18
    Me.CurrentX = 10
    Me.CurrentY = 2
    'form caption
    If cmnDlg.DialogTitle <> "" Then
        Me.Print cmnDlg.DialogTitle
    Else
        Me.Print "Open File"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveMRUs
    SendMessage cdlhwnd, 2, 0, ByVal 0& 'kill commondialog
End Sub

Private Sub Picture1_Resize()
    On Error Resume Next
    SSTab1.Width = Picture1.ScaleWidth + 40
    SSTab1.Height = Picture1.ScaleHeight + 50
    Me.Width = SSTab1.Width * Screen.TwipsPerPixelX + 90
    Me.Height = SSTab1.Height * Screen.TwipsPerPixelY + 420
End Sub

Public Sub LoadRecent()
    'fill listview with recent files and their icons
    Dim z As Long, lItem As ListItem, SFI As SHFILEINFO
    Dim Hico As Long, temp As String
    LV.ListItems.Clear
    Set LV.Icons = Nothing
    Set LV.SmallIcons = Nothing
    ImageListS.ListImages.Clear
    ImageListL.ListImages.Clear
    GetMRUs
    For z = 1 To MRUlist.Count
        If FileExists(MRUlist(z)) Then
            PicTempL.Picture = LoadPicture()
            PicTempS.Picture = LoadPicture()
            PicTempL.Refresh
            PicTempS.Refresh
            'get icons into imagelists
            Hico = SHGetFileInfo(MRUlist(z), 0&, SFI, Len(SFI), SHGFI_ICON Or SHGFI_SMALLICON)
            DrawIconEx PicTempS.hDC, 0, 0, SFI.hIcon, 16, 16, 0, 0, DI_NORMAL
            Hico = SHGetFileInfo(MRUlist(z), 0&, SFI, Len(SFI), SHGFI_ICON Or SHGFI_LARGEICON)
            DrawIconEx PicTempL.hDC, 0, 0, SFI.hIcon, 32, 32, 0, 0, DI_NORMAL
            ImageListL.ListImages.Add , , PicTempL.Image
            ImageListS.ListImages.Add , , PicTempS.Image
            Set lItem = LV.ListItems.Add(, , FileOnly(MRUlist(z)))
            lItem.SubItems(1) = PathOnly(MRUlist(z))
            If Len(lItem.SubItems(1)) < 3 Then lItem.SubItems(1) = lItem.SubItems(1) + "\"
        End If
    Next
    'set imagelists to listview
    If ImageListS.ListImages.Count > 0 Then
        Set LV.Icons = ImageListL
        Set LV.SmallIcons = ImageListS
        For z = 1 To LV.ListItems.Count
            LV.ListItems(z).Icon = z
            LV.ListItems(z).SmallIcon = z
        Next
    End If
End Sub

