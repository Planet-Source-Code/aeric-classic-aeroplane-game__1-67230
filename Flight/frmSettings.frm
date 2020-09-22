VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4080
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPlayerName 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Text            =   "Player"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Turn Name"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton optName 
         Caption         =   "Colour"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optName 
         Caption         =   "Custom"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Start Colour"
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3615
      Begin VB.TextBox txtTurn 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtTurn 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtTurn 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtTurn 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optTurn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Random"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton optTurn 
         BackColor       =   &H00FF0000&
         Caption         =   "Blue"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optTurn 
         BackColor       =   &H0000FF00&
         Caption         =   "Green"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optTurn 
         BackColor       =   &H000000FF&
         Caption         =   "Red"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optTurn 
         BackColor       =   &H0000FFFF&
         Caption         =   "Yellow"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Single Player Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Integer
    If gblnUseName Then optName(1).Value = True Else optName(0).Value = True
    optTurn(gintStartColour).Value = True
     For i = 1 To 4
         txtTurn(i).Text = Player(i)
     Next
    ReadText "Set", 7, PlayerName
    txtPlayerName.Text = PlayerName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
     For i = 1 To 4
          Player(i) = txtTurn(i).Text
     Next
    Me.Hide
End Sub

Private Sub optName_Click(Index As Integer)
Select Case Index
    Case 0
        txtTurn(BLUE).Enabled = False
        txtTurn(GREEN).Enabled = False
        txtTurn(RED).Enabled = False
        txtTurn(YELLOW).Enabled = False
        gblnUseName = False
    Case 1
        txtTurn(BLUE).Enabled = True
        txtTurn(GREEN).Enabled = True
        txtTurn(RED).Enabled = True
        txtTurn(YELLOW).Enabled = True
        gblnUseName = True
End Select
End Sub

Private Sub optTurn_Click(Index As Integer)
    If Index = 0 Then
        blnRandomTurn = True
    Else
        blnRandomTurn = False
        Turn = Index
    End If
    gintStartColour = Index
End Sub
