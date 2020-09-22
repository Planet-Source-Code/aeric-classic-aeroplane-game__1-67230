VERSION 5.00
Begin VB.Form frmRules 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Rules"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   Icon            =   "frmRules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4305
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Rule 4"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   4095
      Begin VB.CheckBox Check4 
         Caption         =   "Roll again when die is 6 (Max 3 times only?)"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Default"
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rule 3"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
      Begin VB.CheckBox Check3 
         Caption         =   "Start only when die is 6"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rule 2"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4095
      Begin VB.CheckBox Check2 
         Caption         =   "Jump if same colour"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rule 1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CheckBox Check1 
         Caption         =   "Kick planes"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This project last modified on 30 November 2006

'''''''Game Logic'''''''''
'0 Load Form
'  Initialize
'1 Click cmdBlueRoll
'  Enable DieTimer
'  Steps generated
'  Disable DieTimer
'2 Check any plane can move
'  Set blnBlue=True to plane that can move
'  Make them enable to click
'3 Enable BlinkTimer
'  Blink enabled planes
'4 If click an enabled plane, blnClick=True
'  Disable BlinkTimer
'  blnClick=False
'  All planes visible
'5 Ask want to go?
'  Yes then enabled FlyTimer
'       If Moves=Steps Then disable FlyTimer, Jump? , next Turn
'       Else moves plane until Moves=Steps
'  No then all blnBlue=False, next Turn
'  Cancel then Enable BlinkTimer
'6 Jump if same colour
'7 Kick other colour plane if occupied

Private Sub Command1_Click()
    SetRules
    Me.Hide
End Sub

Private Sub Command2_Click()
    Check1.Value = 1
    Check2.Value = 1
    Check3.Value = 1
    Check4.Value = 0
End Sub

Private Sub Form_Load()
    If blnKickPlane = False Then Check1.Value = 0 Else Check1.Value = 1
    If blnEnableJump = False Then Check2.Value = 0 Else Check2.Value = 1
    If blnSixOnlyStart = False Then Check3.Value = 0 Else Check3.Value = 1
    If blnRollAgain = False Then Check4.Value = 0 Else Check4.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetRules
    Me.Hide
End Sub

Private Sub SetRules()
    blnKickPlane = Check1.Value
    blnEnableJump = Check2.Value
    blnSixOnlyStart = Check3.Value
    blnRollAgain = Check4.Value
End Sub
