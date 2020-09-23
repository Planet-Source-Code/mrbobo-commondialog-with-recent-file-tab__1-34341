VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CommonDialog Plus"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Open with recent files"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Standard Open"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Standard Save"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

Private Sub Command1_Click()
    'call as you would MS commondialog
    'except instead of .ShowOpen use OpenFile in this example
    With cmnDlg
        .Dorecent = True 'false gives standard dialog
        .Filter = "All files |*.*"
        Set .OwnerFrm = Me
        OpenFile
        MsgBox .FileName
    End With
End Sub

Private Sub Command2_Click()
    With cmnDlg
        .Filter = "All files |*.*"
        Set .OwnerFrm = Me
        OpenFile
        MsgBox .FileName
    End With

End Sub

Private Sub Command3_Click()
    cmnDlg.Filter = "All files |*.*"
    Set cmnDlg.OwnerFrm = Me
    SaveFile
    MsgBox .FileName
End Sub
