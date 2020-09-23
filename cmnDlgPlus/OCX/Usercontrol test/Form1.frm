VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CommonDialog Plus - OCX"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Standard Save"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Standard Open"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin Project1.cmnDlgRecent cmnDlgRecent1 
      Left            =   3000
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open with recent files"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
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
    With cmnDlgRecent1
        .Dorecent = True
        .Filter = "All files |*.*"
        .OwnerFrm = Me
        .ShowOpen
        MsgBox .FileName
    End With
End Sub

Private Sub Command2_Click()
    With cmnDlgRecent1
        .Filter = "All files |*.*"
        .OwnerFrm = Me
        .ShowOpen
        MsgBox .FileName
    End With
End Sub

Private Sub Command3_Click()
    With cmnDlgRecent1
        .Filter = "All files |*.*"
        .OwnerFrm = Me
        .ShowSave
        MsgBox .FileName
    End With

End Sub
