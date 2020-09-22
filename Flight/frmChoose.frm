VERSION 5.00
Begin VB.Form frmChoose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose your colour"
   ClientHeight    =   1560
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3285
   Icon            =   "frmChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.PictureBox picPlane 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   3
      Left            =   1800
      Picture         =   "frmChoose.frx":08CA
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   240
      Width           =   390
   End
   Begin VB.PictureBox picPlane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   4
      Left            =   2520
      Picture         =   "frmChoose.frx":0D46
      ScaleHeight     =   390
      ScaleWidth      =   390
      TabIndex        =   2
      Top             =   240
      Width           =   420
   End
   Begin VB.PictureBox picPlane 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   2
      Left            =   1080
      Picture         =   "frmChoose.frx":11BD
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   1
      Top             =   240
      Width           =   390
   End
   Begin VB.PictureBox picPlane 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   360
      Picture         =   "frmChoose.frx":163E
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   240
      Width           =   390
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is created on 23 October 2004 1:53 AM
'Use Keyboard Arrow to select Plane Colour?
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    picPlane(PlayerColour).BackColor = vbBlack
End Sub

Private Sub picPlane_Click(Index As Integer)
    picPlane(PlayerColour).BackColor = vbWhite
    blnAI(PlayerColour) = True
    PlayerColour = Index
    blnAI(PlayerColour) = False
    picPlane(PlayerColour).BackColor = vbBlack
End Sub
