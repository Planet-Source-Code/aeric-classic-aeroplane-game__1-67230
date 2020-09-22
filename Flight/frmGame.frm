VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Flight"
   ClientHeight    =   8175
   ClientLeft      =   1530
   ClientTop       =   750
   ClientWidth     =   8175
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Frame fraMessage 
         BackColor       =   &H00FF0000&
         Height          =   1455
         Left            =   2760
         TabIndex        =   23
         Top             =   3360
         Visible         =   0   'False
         Width           =   2655
         Begin VB.Label lblMessage 
            Alignment       =   2  'Center
            Caption         =   "Waiting Blue (AI)..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1095
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Timer AIThinkTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4560
         Top             =   4920
      End
      Begin VB.Timer DelayTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4560
         Top             =   2880
      End
      Begin VB.Timer ShortCutTimer 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   6480
         Top             =   4560
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Text            =   "6"
         ToolTipText     =   "1 to 6"
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   14
         Left            =   3000
         Picture         =   "frmGame.frx":0442
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   11
         Left            =   3000
         Picture         =   "frmGame.frx":08BE
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   12
         Left            =   3360
         Picture         =   "frmGame.frx":0D37
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   13
         Left            =   3360
         Picture         =   "frmGame.frx":11B4
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdRoll 
         Caption         =   "Roll"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdRoll 
         Caption         =   "Roll"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   6960
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdRoll 
         Caption         =   "Roll"
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   720
         TabIndex        =   18
         Top             =   6960
         Width           =   495
      End
      Begin VB.CommandButton cmdRoll 
         Caption         =   "Roll"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   6960
         TabIndex        =   17
         Top             =   6960
         Width           =   495
      End
      Begin VB.Timer DieTimer 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1200
         Top             =   3120
      End
      Begin VB.Timer FlyTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6480
         Top             =   3120
      End
      Begin VB.Timer BlinkTimer 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1200
         Top             =   4680
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   21
         Left            =   4440
         Picture         =   "frmGame.frx":1630
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   24
         Left            =   4440
         Picture         =   "frmGame.frx":1AB6
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   22
         Left            =   4800
         Picture         =   "frmGame.frx":1F37
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   23
         Left            =   4800
         Picture         =   "frmGame.frx":23B9
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         Height          =   375
         Index           =   41
         Left            =   3000
         Picture         =   "frmGame.frx":283A
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   8
         Top             =   6360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picPlane 
         Height          =   375
         Index           =   44
         Left            =   3000
         Picture         =   "frmGame.frx":2CB5
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   6720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picPlane 
         Height          =   375
         Index           =   42
         Left            =   3360
         Picture         =   "frmGame.frx":3133
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   6
         Top             =   6360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picPlane 
         Height          =   375
         Index           =   43
         Left            =   3360
         Picture         =   "frmGame.frx":35B1
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   5
         Top             =   6720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   31
         Left            =   4440
         Picture         =   "frmGame.frx":3A28
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   4
         Top             =   6360
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   34
         Left            =   4440
         Picture         =   "frmGame.frx":3EA0
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   3
         Top             =   6720
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   32
         Left            =   4800
         Picture         =   "frmGame.frx":4322
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   2
         Top             =   6360
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picPlane 
         AutoSize        =   -1  'True
         Height          =   420
         Index           =   33
         Left            =   4800
         Picture         =   "frmGame.frx":47A6
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   1
         Top             =   6720
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   11
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   12
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   13
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   5
         Left            =   40
         Shape           =   3  'Circle
         Top             =   1960
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   31
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   32
         Left            =   6480
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   33
         Left            =   6480
         Stretch         =   -1  'True
         Top             =   7440
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   41
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   7440
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   42
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   43
         Left            =   240
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   495
      End
      Begin VB.Image Plane 
         Height          =   495
         Index           =   21
         Left            =   6480
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Image Plane 
         Height          =   495
         Index           =   22
         Left            =   6480
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Plane 
         Height          =   495
         Index           =   23
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Plane 
         Height          =   495
         Index           =   24
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   79
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   5280
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   44
         Left            =   240
         Stretch         =   -1  'True
         Top             =   7440
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   34
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   11
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   12
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Plane 
         Enabled         =   0   'False
         Height          =   495
         Index           =   14
         Left            =   240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   14
         Left            =   240
         Shape           =   3  'Circle
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   13
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   21
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   31
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   6480
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00008080&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   43
         Left            =   240
         Shape           =   3  'Circle
         Top             =   6480
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   32
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   6480
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   33
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   34
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   24
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   22
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   23
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00008080&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   42
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   6480
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00008080&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   44
         Left            =   240
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Base 
         BackColor       =   &H00008080&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   41
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Base3 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1935
         Left            =   6240
         Top             =   6240
         Width           =   1935
      End
      Begin VB.Shape Path 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   7
         Left            =   7640
         Shape           =   3  'Circle
         Top             =   5720
         Width           =   495
      End
      Begin VB.Shape Base1 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1935
         Left            =   0
         Top             =   0
         Width           =   1935
      End
      Begin VB.Shape Base4 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1935
         Left            =   0
         Top             =   6240
         Width           =   1935
      End
      Begin VB.Shape Path 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   8
         Left            =   1960
         Shape           =   3  'Circle
         Top             =   7640
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   9
         Left            =   420
         Shape           =   3  'Circle
         Top             =   2360
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   24
         Left            =   5520
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   23
         Left            =   5520
         Shape           =   3  'Circle
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   13
         Left            =   2360
         Shape           =   3  'Circle
         Top             =   1960
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   17
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   21
         Left            =   4800
         Shape           =   3  'Circle
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   11
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   12
         Left            =   1960
         Shape           =   3  'Circle
         Top             =   2320
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   14
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   15
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   16
         Left            =   2340
         Shape           =   3  'Circle
         Top             =   420
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   19
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   20
         Left            =   4320
         Shape           =   3  'Circle
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   25
         Left            =   5320
         Shape           =   3  'Circle
         Top             =   1960
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   10
         Left            =   960
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   18
         Left            =   3360
         Shape           =   3  'Circle
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   22
         Left            =   5320
         Shape           =   3  'Circle
         Top             =   420
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   26
         Left            =   5720
         Shape           =   3  'Circle
         Top             =   2340
         Width           =   495
      End
      Begin VB.Shape GreenPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   19
         Left            =   2880
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   10
         Left            =   4800
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   13
         Left            =   1920
         Top             =   1440
         Width           =   975
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   11
         Left            =   960
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   6
         Left            =   1440
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   8
         Left            =   1920
         Top             =   960
         Width           =   975
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   14
         Left            =   3360
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape GreenPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   20
         Left            =   3840
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   15
         Left            =   4320
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   27
         Left            =   6240
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   60
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2880
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   59
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   58
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   57
         Left            =   240
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   56
         Left            =   240
         Shape           =   3  'Circle
         Top             =   4800
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   55
         Left            =   420
         Shape           =   3  'Circle
         Top             =   5320
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   54
         Left            =   960
         Shape           =   3  'Circle
         Top             =   5520
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   53
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   5520
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   52
         Left            =   1960
         Shape           =   3  'Circle
         Top             =   5320
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   51
         Left            =   2360
         Shape           =   3  'Circle
         Top             =   5720
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   50
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   49
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   6720
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   48
         Left            =   2360
         Shape           =   3  'Circle
         Top             =   7240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   47
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   46
         Left            =   3360
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   45
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   44
         Left            =   4320
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   43
         Left            =   4800
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   42
         Left            =   5320
         Shape           =   3  'Circle
         Top             =   7240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   41
         Left            =   5520
         Shape           =   3  'Circle
         Top             =   6720
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   40
         Left            =   5520
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   39
         Left            =   5325
         Shape           =   3  'Circle
         Top             =   5720
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   38
         Left            =   5720
         Shape           =   3  'Circle
         Top             =   5320
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   37
         Left            =   6240
         Shape           =   3  'Circle
         Top             =   5520
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   36
         Left            =   6720
         Shape           =   3  'Circle
         Top             =   5520
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   35
         Left            =   7240
         Shape           =   3  'Circle
         Top             =   5320
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   34
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   4800
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   33
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   32
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   31
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   30
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   2880
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   29
         Left            =   7240
         Shape           =   3  'Circle
         Top             =   2360
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   28
         Left            =   6720
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape BluePath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   20
         Left            =   0
         Top             =   3840
         Width           =   975
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   5
         Left            =   0
         Top             =   3360
         Width           =   975
      End
      Begin VB.Shape BluePath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   26
         Left            =   0
         Top             =   2880
         Width           =   975
      End
      Begin VB.Shape GreenPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   4
         Left            =   7200
         Top             =   3360
         Width           =   975
      End
      Begin VB.Shape BluePath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   9
         Left            =   7200
         Top             =   2880
         Width           =   975
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   14
         Left            =   0
         Top             =   4320
         Width           =   975
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   6
         Left            =   0
         Top             =   4800
         Width           =   975
      End
      Begin VB.Shape BluePath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   14
         Left            =   960
         Top             =   5280
         Width           =   495
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   19
         Left            =   1440
         Top             =   5280
         Width           =   495
      End
      Begin VB.Shape GreenPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   10
         Left            =   2880
         Top             =   7200
         Width           =   495
      End
      Begin VB.Shape BluePath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   15
         Left            =   3360
         Top             =   7200
         Width           =   495
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   20
         Left            =   3840
         Top             =   7200
         Width           =   495
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   4
         Left            =   4320
         Top             =   7200
         Width           =   495
      End
      Begin VB.Shape GreenPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   9
         Left            =   4800
         Top             =   7200
         Width           =   495
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   18
         Left            =   5280
         Top             =   6720
         Width           =   975
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   2
         Left            =   5280
         Top             =   6240
         Width           =   975
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   17
         Left            =   6240
         Top             =   5280
         Width           =   495
      End
      Begin VB.Shape BluePath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   10
         Left            =   7200
         Top             =   4800
         Width           =   975
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   15
         Left            =   7200
         Top             =   4320
         Width           =   975
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   1
         Left            =   6720
         Top             =   5280
         Width           =   495
      End
      Begin VB.Shape GreenPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   8
         Left            =   1920
         Top             =   6720
         Width           =   975
      End
      Begin VB.Shape BluePath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   13
         Left            =   1920
         Top             =   6240
         Width           =   975
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   20
         Left            =   7200
         Top             =   3840
         Width           =   975
      End
      Begin VB.Shape Path 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   6
         Left            =   5720
         Shape           =   3  'Circle
         Top             =   40
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   80
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   4800
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   78
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   5760
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   77
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   76
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   6720
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   75
         Left            =   4800
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   74
         Left            =   5280
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   73
         Left            =   5760
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   72
         Left            =   6240
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   71
         Left            =   6720
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   70
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   2880
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   69
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   68
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   67
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   66
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   65
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   64
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   63
         Left            =   1920
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   62
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Path 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   61
         Left            =   960
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape GreenPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   2415
         Index           =   23
         Left            =   3840
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   23
         Left            =   4800
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         Height          =   2415
         Index           =   23
         Left            =   3840
         Top             =   4800
         Width           =   495
      End
      Begin VB.Shape BluePath 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   21
         Left            =   960
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Shape Goal 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   4
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   495
      End
      Begin VB.Shape Goal 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   2
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   495
      End
      Begin VB.Shape Goal 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   3
         Left            =   4320
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Shape Goal 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   1
         Left            =   3360
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   0
         Picture         =   "frmGame.frx":4C22
         Top             =   5280
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   975
         Left            =   1920
         Picture         =   "frmGame.frx":5154
         Top             =   5280
         Width           =   975
      End
      Begin VB.Shape GreenFly 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFC0&
         Height          =   495
         Index           =   1
         Left            =   4320
         Top             =   5720
         Width           =   495
      End
      Begin VB.Shape GreenFly 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFC0&
         Height          =   495
         Index           =   2
         Left            =   3360
         Top             =   5720
         Width           =   495
      End
      Begin VB.Shape GreenFly 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFC0&
         Height          =   495
         Index           =   3
         Left            =   2880
         Top             =   5720
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   960
         Left            =   0
         Picture         =   "frmGame.frx":5679
         Top             =   1920
         Width           =   960
      End
      Begin VB.Image Image4 
         Height          =   975
         Left            =   1920
         Picture         =   "frmGame.frx":5B7B
         Top             =   0
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   975
         Left            =   5280
         Picture         =   "frmGame.frx":607C
         Top             =   7200
         Width           =   975
      End
      Begin VB.Image Image12 
         Height          =   975
         Left            =   7200
         Picture         =   "frmGame.frx":657D
         Top             =   5280
         Width           =   975
      End
      Begin VB.Shape YellowFly 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFFF&
         Height          =   495
         Index           =   1
         Left            =   3360
         Top             =   1960
         Width           =   495
      End
      Begin VB.Shape YellowFly 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFFF&
         Height          =   495
         Index           =   2
         Left            =   4320
         Top             =   1960
         Width           =   495
      End
      Begin VB.Shape RedFly 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0FF&
         Height          =   495
         Index           =   2
         Left            =   1960
         Top             =   3360
         Width           =   495
      End
      Begin VB.Shape RedFly 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0FF&
         Height          =   495
         Index           =   7
         Left            =   1960
         Top             =   4320
         Width           =   495
      End
      Begin VB.Shape BlueFly 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFC0C0&
         Height          =   495
         Index           =   1
         Left            =   5720
         Top             =   3360
         Width           =   495
      End
      Begin VB.Shape BlueFly 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFC0C0&
         Height          =   495
         Index           =   2
         Left            =   5720
         Top             =   4320
         Width           =   495
      End
      Begin VB.Image Image10 
         Height          =   975
         Left            =   5280
         Picture         =   "frmGame.frx":6A7E
         Top             =   5280
         Width           =   975
      End
      Begin VB.Shape BlueFly 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFC0C0&
         Height          =   495
         Index           =   3
         Left            =   5720
         Top             =   4800
         Width           =   495
      End
      Begin VB.Shape RedFly 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0FF&
         Height          =   495
         Index           =   6
         Left            =   1960
         Top             =   4800
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   975
         Left            =   1920
         Picture         =   "frmGame.frx":6F9C
         Top             =   1920
         Width           =   975
      End
      Begin VB.Shape YellowFly 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFFF&
         Height          =   495
         Index           =   0
         Left            =   2880
         Top             =   1960
         Width           =   495
      End
      Begin VB.Shape GreenFly 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFC0&
         Height          =   495
         Index           =   0
         Left            =   4800
         Top             =   5720
         Width           =   495
      End
      Begin VB.Shape RedFly 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0FF&
         Height          =   495
         Index           =   3
         Left            =   1960
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image Image11 
         Height          =   975
         Left            =   1920
         Picture         =   "frmGame.frx":74BA
         Top             =   7200
         Width           =   975
      End
      Begin VB.Shape GreenPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   2
         Left            =   6240
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape RedPath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   18
         Left            =   6720
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape Base2 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1935
         Left            =   6240
         Top             =   0
         Width           =   1935
      End
      Begin VB.Image Image6 
         Height          =   975
         Left            =   5280
         Picture         =   "frmGame.frx":79AC
         Top             =   0
         Width           =   975
      End
      Begin VB.Shape YellowPath 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   11
         Left            =   5280
         Top             =   960
         Width           =   975
      End
      Begin VB.Shape BluePath 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   6
         Left            =   5280
         Top             =   1440
         Width           =   975
      End
      Begin VB.Image Image7 
         Height          =   975
         Left            =   5280
         Picture         =   "frmGame.frx":7E9E
         Top             =   1920
         Width           =   975
      End
      Begin VB.Shape YellowFly 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFFF&
         Height          =   495
         Index           =   3
         Left            =   4800
         Top             =   1960
         Width           =   495
      End
      Begin VB.Image Image8 
         Height          =   975
         Left            =   7200
         Picture         =   "frmGame.frx":83C0
         Top             =   1920
         Width           =   975
      End
      Begin VB.Shape BlueFly 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFC0C0&
         Height          =   495
         Index           =   0
         Left            =   5720
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image Image13 
         Height          =   1455
         Left            =   3360
         Picture         =   "frmGame.frx":88B2
         Top             =   3360
         Width           =   1455
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Game"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start Game"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "&End Game"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOne 
         Caption         =   "&One Player"
         Checked         =   -1  'True
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFour 
         Caption         =   "&Four Players"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSim 
         Caption         =   "&Simulation (4 AIs)"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Option"
      Begin VB.Menu mnuMenu 
         Caption         =   "Hide &Menu (Esc)"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuRules 
         Caption         =   "Game &Rules"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Enable Fixed &Die"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Seed As Integer
Private RollCount As Integer
Private Moves As Integer
Private Blink As Integer
Private ThinkTime As Integer
Private Delay As Integer

Private blnClick As Boolean 'True = clicked a plane
Private Go As Integer 'User clicked/Moving Plane Index
Private StopPath As Integer 'Path Number of currently/last moving plane
Private intFlying As Integer 'For ShortCut Flying
Private intStartClick As Integer 'First round player click
Private intMaxRoll As Integer ' Maximum times to roll a die in 1 turn
Private intRepeatRoll As Integer ' How many times a die roll repeatedly

Private COLOUR(1 To 4, 1 To 4) As Integer 'Plane position (Colour, Number)
Private blnColour(1 To 4, 1 To 4) As Boolean 'Plane Ready? (Colour, Number) for blinking
Private IMG(1 To 4, 1 To 4) As Integer 'Load Which Image? (Colour, Number)
Private Win(1 To 4) As Integer 'How many plane goal
Private blnVictory(1 To 4) As Boolean  'Victorious?
Private blnReverse(1 To 4, 1 To 4) As Boolean 'fly reverse direction

Private ShowMenu As Boolean

Private OhOu As String 'Sound path 'Use Static or Constant?
Private Play As String 'Sound path

'Private Sub cmdBlueRoll_KeyDown(KeyCode As Integer, Shift As Integer)
'18 = vbKeyMenu (Alt)
'19 = vbKeyPause
'27 = vbKeyEscape
'If KeyCode = 18 Or KeyCode = 19 Or KeyCode = 27 Then
'If Me.Caption = "Flight Game (Testing only)" Then
'    Me.Caption = Me.Caption & " - Paused (Press Esc to Resume)"
'Else
'    Me.Caption = "Flight Game (Testing only)"
'End If
'mnuMenu_Click
'End If
'End Sub

Private Sub cmdRoll_Click(Index As Integer)
    'Roll one time only
    'Human cannot click or roll die more than once
    'set boolean = true if clicked and do not allowed human click
    If Turn = Index Then 'And blnRoll = False Then
'        blnRoll = True
'        Rollable = True
        Randomize
        RollCount = Rnd() * 100 Mod 6 + 6
        DieTimer.Enabled = True
        cmdRoll(Turn).Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    Dim dv As String 'debug value
    'Debug purpose
    RollCount = 0
    dv = Text1
    If Not (dv = "1" Or dv = "2" Or dv = "3" Or dv = "4" Or dv = "5" Or dv = "6") Then Exit Sub
    Steps = dv
    cmdRoll(Turn).Caption = Steps
    BlinkTimer.Enabled = True
    cmdRoll(Turn).Enabled = False
    EnablePlane
    'Text1.Enabled = False
    'Command1.Enabled = False
End Sub

Private Sub DieTimer_Timer()
    DieNumber = (Rnd() * 10 + Seed) Mod 6
    Steps = DieNumber + 1
    cmdRoll(Turn).Caption = Steps
    If RollCount = 0 Then
        DieTimer.Enabled = False
        
        If intStartClick < 1 Then ' Or intStartClick = 0
            blnFirstRound = False
            'Enable Load/Save Game after a game is loaded
            mnuLoad.Enabled = True
            mnuSave.Enabled = True
        Else
            blnFirstRound = True
            intStartClick = intStartClick - 1
        End If
        
        BlinkTimer.Enabled = True 'Blink AI moveable planes
        
        'At the same time let AI thinks
        If blnAI(Turn) = True Then
            Go = AI(AINum(Turn))
            If Go > 0 Then
                Delay = 0 'Thinking Delay
                AIThinkTimer.Enabled = True
            End If
        End If
        Exit Sub
    End If
    RollCount = RollCount - 1
End Sub

Private Sub Form_Load()
Dim Dummy As Integer
On Error GoTo err

Me.Top = -100

'Default Game Mode
GameMode = "A"
PlayerColour = BLUE
Turn = PlayerColour 'Default Turn
LoadSetting
LoadRule

Open App.Path & "\Saved.txt" For Input As #3
    If Not EOF(3) Then
        Input #3, Player(BLUE), Dummy, Dummy, Dummy, Dummy, Dummy
        Input #3, Player(GREEN), Dummy, Dummy, Dummy, Dummy, Dummy
        Input #3, Player(RED), Dummy, Dummy, Dummy, Dummy, Dummy
        Input #3, Player(YELLOW), Dummy, Dummy, Dummy, Dummy, Dummy
    End If
Close #3
    
InitPlane

'This should be disabled
mnuFile.Visible = True 'False
mnuGame.Visible = True 'False
mnuOption.Visible = True 'False
mnuHelp.Visible = True 'False
mnuMenu.Checked = False 'True
ShowMenu = True

'Generate Seed
Randomize
Seed = ((Rnd() * 100) + (Timer * 100)) Mod 6
If Seed = 0 Then Seed = 1

intStartClick = 4
blnFirstRound = True

'Maximum roll repeat if die number is 6
intMaxRoll = 2

'Play sound
OhOu = App.Path & "\Oh Ou.wav"
Play = App.Path & "\Play.wav"

Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & " (built " & App.Revision & ")"
Exit Sub
err:
    AppendText "Error", Now & " " & " Error: " & Error & "(Form Load)"
End Sub

Private Function InitPlane()
    Dim i As Integer
    
    For i = 1 To 4 'BLUE To YELLOW
        'Initial positions
        COLOUR(BLUE, i) = BLUE
        COLOUR(GREEN, i) = GREEN
        COLOUR(RED, i) = RED
        COLOUR(YELLOW, i) = YELLOW
        
        ' Use For Loop later
        ' Eg. For i = 1
        ' IMG(BLUE, 1) = PathDirection(1)   -> means Blue1 = South
        ' Plane(11).Picture = picPlane(1, IMG(BLUE, 1)).Picture
        IMG(BLUE, i) = PathDirection(COLOUR(BLUE, i))
        Plane(BLUE * 10 + i).Picture = picPlane(BLUE * 10 + BLUE).Picture
        IMG(GREEN, i) = PathDirection(COLOUR(GREEN, i))
        Plane(GREEN * 10 + i).Picture = picPlane(GREEN * 10 + GREEN).Picture
        IMG(RED, i) = PathDirection(COLOUR(RED, i))
        Plane(RED * 10 + i).Picture = picPlane(RED * 10 + RED).Picture
        IMG(YELLOW, i) = PathDirection(COLOUR(YELLOW, i))
        Plane(YELLOW * 10 + i).Picture = picPlane(YELLOW * 10 + YELLOW).Picture
        
        blnColour(BLUE, i) = False
        blnColour(GREEN, i) = False
        blnColour(RED, i) = False
        blnColour(YELLOW, i) = False
        
        ' blnAI(i) = True
        blnVictory(i) = False
        Win(i) = 0
    Next

    'Generate AI
    If PlayerColour <> BLUE Then AINum(BLUE) = GenerateAI
    If PlayerColour <> GREEN Then AINum(GREEN) = GenerateAI
    If PlayerColour <> RED Then AINum(RED) = GenerateAI
    If PlayerColour <> YELLOW Then AINum(YELLOW) = GenerateAI

End Function

Private Sub FlyTimer_Timer()
Dim cut As Integer
Dim FlyPlane As Integer
Dim PlaneNo As Integer
Dim D1 As Integer
Dim D2 As Integer
Dim GoWin As Integer
Dim Jump As Integer

'On Error GoTo errFly

FlyPlane = COLOUR(Turn, Go)
PlaneNo = Turn * 10

'Make Plane on top
Plane(PlaneNo + Go).ZOrder 0

    Select Case Turn
        Case BLUE
            D1 = 58
            D2 = 66
            GoWin = 4
            Jump = 2
        Case GREEN
            D1 = 19
            D2 = 71
            GoWin = 1
            Jump = 3
        Case RED
            D1 = 32
            D2 = 76
            GoWin = 2
            Jump = 0
        Case YELLOW
            D1 = 45
            D2 = 81
            GoWin = 3
            Jump = 1
    End Select

    'Inside Base
    If FlyPlane = Turn Then
        'Direction
        FlyPlane = FlyPlane + 4 ' Or Turn + 4
        COLOUR(Turn, Go) = FlyPlane
        StopPath = FlyPlane
        Plane(PlaneNo + Go).Left = Path(FlyPlane).Left
        Plane(PlaneNo + Go).Top = Path(FlyPlane).Top
        FlyTimer.Enabled = False 'Disable this line cause Green moves
'        NextTurn 'Cannot repeat roll die
            If Steps = 6 And blnRollAgain = True Then  'Roll again if six
                If intRepeatRoll < intMaxRoll Then
                    SameTurn
                    'intRepeatRoll = intRepeatRoll + 1
                Else
                    NextTurn ' No repeat
                End If
            Else
                NextTurn ' No repeat
            End If
        Exit Sub
    End If

    If Moves = Steps Then ' use >= ? to prevent undefinite loop
        FlyTimer.Enabled = False
        blnReverse(Turn, Go) = False
        If FlyPlane = D2 Then 'Win
            ' Plane(PlaneNo + Go).Left = Goal(Turn).Left 'Need?
            ' Plane(PlaneNo + Go).Top = Goal(Turn).Top 'Need?
            FlyPlane = 0 'Cannot move anymore
            COLOUR(Turn, Go) = FlyPlane
            StopPath = FlyPlane
            Win(Turn) = Win(Turn) + 1
            If Win(Turn) = 4 Then
                If GameMode <> "S" Then 'Avoid Messagebox
                    MsgBox Player(Turn) & "'s " & gstrObject & " has completed the mission.", vbOKOnly, "Mission accomplished"
                End If
                'no more turn
                blnVictory(Turn) = True
            Else
                If GameMode <> "S" Then
                    MsgBox Player(Turn) & "'s " & gstrObject & " " & Go & " has completed objective " & Win(Turn), vbOKOnly, "Objective complete"
                End If
            End If
            'NextTurn
            If Steps = 6 And blnRollAgain = True Then  'Roll again if six
                If intRepeatRoll < intMaxRoll Then
                    SameTurn
                    'intRepeatRoll = intRepeatRoll + 1
                Else
                    NextTurn ' No repeat
                End If
            Else
                NextTurn ' No repeat
            End If
            Exit Sub
        End If
        
        'If COLOUR(1,Go)=26, Take ShortCut
        cut = ((Turn Mod 4) + 1) * 13
        If FlyPlane = cut Then
            'FlyTimer.Enabled = False
            IMG(Turn, Go) = PathDirection(FlyPlane)
            Plane(PlaneNo + Go).Picture = picPlane(PlaneNo + Turn).Picture
            intFlying = Path(cut).Left
            intFlying = Path(cut).Top
            ShortCutTimer.Enabled = True
            'MsgBox "Before taking shortcut"
            Exit Sub
        End If
        
        'If COLOUR(BLUE,Go)=58 or COLOUR(BLUE,Go)=62, don't jump!
        'Else same colour then jump 1 time
        If Not (FlyPlane = D1 Or FlyPlane = D2 - 4) Then
            If blnEnableJump = True And FlyPlane Mod 4 = Jump Then
                FlyPlane = FlyPlane + 4 'Do not use FlyPlane
                ' continue next section after jump
                ' dun let blue jumps
                If FlyPlane = 61 Then FlyPlane = 9 ' Yellow
                If FlyPlane = 63 Then FlyPlane = 11 ' Green
                If FlyPlane = 64 Then FlyPlane = 12 ' Red
                COLOUR(Turn, Go) = FlyPlane
                Direction
                Plane(PlaneNo + Go).Left = Path(FlyPlane).Left
                Plane(PlaneNo + Go).Top = Path(FlyPlane).Top
            End If
        End If '????
        
        StopPath = FlyPlane
        Direction 'Direction
        If blnKickPlane = True Then
            CheckKick
        End If
        ' NextTurn
        If Steps = 6 And blnRollAgain = True Then  'Roll again if six
            If intRepeatRoll < intMaxRoll Then
                SameTurn
                'intRepeatRoll = intRepeatRoll + 1
            Else
                NextTurn ' No repeat
            End If
        Else
            NextTurn ' No repeat
        End If
        Exit Sub
    End If
    
    'If still got available Moves
    'Turn Plane before flying ?
    If FlyPlane = Turn + 4 Then
    'In between Base and Circuit
        FlyPlane = D1 + 3
        If FlyPlane = 61 Then FlyPlane = 9 ' if Blue
        COLOUR(Turn, Go) = FlyPlane
        Plane(PlaneNo + Go).Left = Path(FlyPlane).Left
        Plane(PlaneNo + Go).Top = Path(FlyPlane).Top
    Else
        If blnReverse(Turn, Go) = False Then
            FlyPlane = FlyPlane + 1
        Else
            FlyPlane = FlyPlane - 1
        End If
        If FlyPlane = D1 + 1 Then '59
            FlyPlane = D2 - 5 '61 'Go towards goal
        Else
            If FlyPlane = 61 Then
                FlyPlane = 9 'Continue next section (except Blue)
            End If
        End If
        COLOUR(Turn, Go) = FlyPlane
        If FlyPlane = D2 Then '66
            Plane(PlaneNo + Go).Left = Goal(Turn).Left
            Plane(PlaneNo + Go).Top = Goal(Turn).Top
            Direction
            blnReverse(Turn, Go) = True
        Else
            Direction
            Plane(PlaneNo + Go).Left = Path(FlyPlane).Left
            Plane(PlaneNo + Go).Top = Path(FlyPlane).Top
        End If
    End If
    
    If Moves < Steps Then
        Moves = Moves + 1
        Direction
    End If
Exit Sub
errFly:
    MsgBox Error & Chr(13) & gstrObject & " (" & Player(Turn) & ") = " & Go, vbExclamation, "Flytimer"
    FlyTimer.Enabled = False
End Sub

'Use for Human and AI Players
'Check whether any plane enable?
Private Function CheckMove()
    Dim i As Integer
    
    For i = 1 To 4
        If COLOUR(Turn, i) < 5 Then
            If COLOUR(Turn, i) = 0 Then
                blnColour(Turn, i) = False
            Else
                If Steps = 6 Then
                    blnColour(Turn, i) = True
                Else '==========DieNumber Not 6==========
                    If blnFirstRound = True And blnSixOnlyStart = False Then
                        blnColour(Turn, i) = True
                    Else
                        blnColour(Turn, i) = False
                    End If
                End If
            End If
        Else
            blnColour(Turn, i) = True
        End If
        
    Next
'blink plane
'If Turn <> PlayerColour Then
'    BlinkTimer.Enabled = True
'End If
End Function

Private Sub BlinkTimer_Timer()
Dim i As Integer
    
    'Next turn if no selected plane for AI
    If blnAI(Turn) And Go = 0 Then
        BlinkTimer.Enabled = False
        AIThinkTimer.Enabled = False 'optional?
        'blnClick = False
        NextTurn 'No repeat?
        Exit Sub
    End If
    
    ' Check whether plane is enabled before blink
    CheckMove
    
    'Next turn if no enabled planes
    If blnColour(Turn, 1) = False And blnColour(Turn, 2) = False And blnColour(Turn, 3) = False And blnColour(Turn, 4) = False Then
        BlinkTimer.Enabled = False
        AIThinkTimer.Enabled = False
        'blnClick = False 'Require?
        NextTurn 'No repeat?
        Exit Sub
    End If
        
    If blnAI(Turn) = False Then
        fraMessage.Visible = False
        EnablePlane ' Must CheckMove first
    End If
    
    For i = 1 To 4
        If Blink = 0 Then 'Blink Plane that ready to move
            If blnColour(Turn, i) = True Then Plane(Turn * 10 + i).Picture = LoadPicture("")
        Else
            If blnColour(Turn, i) = True Then Plane(Turn * 10 + i).Picture = picPlane(Turn * 10 + IMG(Turn, i)).Picture
        End If
    Next
    
    If blnClick = True Then
        'Make all planes visible
        For i = 1 To 4
            Plane(Turn * 10 + i).Picture = picPlane(Turn * 10 + IMG(Turn, i)).Picture
            blnColour(Turn, i) = False 'disabled all planes
        Next
        If blnAI(Turn) = True Then 'If this is an AI then don't ask
            'Check Go = ?
            BlinkTimer.Enabled = False 'mandatory
            'AIThinkTimer.Enabled = False 'optional
            blnClick = False 'require
            FlyTimer.Enabled = True 'mandatory
        Else
'            EnablePlane ' Must CheckMove first
            ' Ask Player
            If vbCancel = MsgBox("Move " & gstrObject & " " & Go & " ?", vbOKCancel + vbQuestion, Player(Turn) & "'s Turn") Then
                'Cancel selected plane
                'Check available planes again
'                CheckMove
                'Keep blinking enabled plane(s)
                ' BlinkTimer.Enabled = True
'                fraMessage.Visible = True
                blnClick = False
            Else
                BlinkTimer.Enabled = False
                'AIThinkTimer.Enabled = False
                blnClick = False 'require
                FlyTimer.Enabled = True
                DisablePlane
            End If
        End If
'        blnClick = False
    End If

    Blink = Blink + 1
    If Blink = 2 Then Blink = 0
End Sub

Private Sub Form_Resize()
'If Me.WindowState = 0 Then
'picBoard.Top = picBoard.Top - 120
'End If
picBoard.Left = (Me.Width - picBoard.Width) / 2 - 60
picBoard.Top = (Me.Height - picBoard.Height) / 2 - 280
If Me.WindowState = 0 Then
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
End If
End Sub

Private Sub mnuLoad_Click()
    LoadGame
    SetPlanePosition
End Sub

Private Sub mnuSave_Click()
    SaveGame
End Sub

Private Sub mnuSim_Click()
    'Dim i As Integer
    mnuOne.Checked = False
    mnuFour.Checked = False
    mnuSim.Checked = True
    GameMode = "S"
    'For i = BLUE To YELLOW
    '    blnAI(i) = True
    'Next
    DelayTimer.Interval = 500
    FlyTimer.Interval = 500
    AIThinkTimer.Interval = 500
    intStartClick = 4
End Sub

Private Sub Plane_Click(Index As Integer)
    If blnColour(Index / 10, Index Mod 10) = True And blnAI(Turn) = False Then
        Go = Index Mod 10
        blnClick = True
        'lblMessage.Caption = ""
        'fraMessage.Visible = False
        'Make Plane on top (see FlyTimer)
        ' Plane(Index).ZOrder 0
    End If
End Sub

Private Sub mnuEnd_Click()
    'Disable All Timer objects
    DieTimer.Enabled = False
    BlinkTimer.Enabled = False
    FlyTimer.Enabled = False
    ShortCutTimer.Enabled = False
    DelayTimer.Enabled = False
    AIThinkTimer.Enabled = False
    Unload Me
    frmGame.Show
End Sub

Private Sub mnuOne_Click()
    mnuOne.Checked = True
    mnuFour.Checked = False
    mnuSim.Checked = False
    frmChoose.Show vbModal
    GameMode = "A"
    DelayTimer.Interval = 1000
    FlyTimer.Interval = 1000
    AIThinkTimer.Interval = 1000
    intStartClick = 4
End Sub

Private Sub mnuFour_Click()
    'Dim i As Integer
    mnuOne.Checked = False
    mnuFour.Checked = True
    mnuSim.Checked = False
    GameMode = "B"
    'For i = BLUE To YELLOW
    '    blnAI(i) = False
    'Next
    DelayTimer.Interval = 1000
    FlyTimer.Interval = 1000
    AIThinkTimer.Interval = 1000
    intStartClick = 4
End Sub

Private Sub mnuSettings_Click()
    frmSettings.Show vbModal
End Sub

Private Sub mnuStart_Click()
Dim iCol As Integer

    mnuStart.Enabled = False
    mnuOne.Enabled = False
    mnuFour.Enabled = False
    mnuSim.Enabled = False
    mnuSettings.Enabled = False
    mnuRules.Enabled = False
    mnuEnd.Enabled = True

    'Set this in Reset game too
    cmdRoll(BLUE).Enabled = False
    cmdRoll(GREEN).Enabled = False
    cmdRoll(RED).Enabled = False
    cmdRoll(YELLOW).Enabled = False
    
    ' Enable Load/Save game after first round
    If intStartClick = 0 Then
        'Enable Load/Save Game after a game is loaded
        mnuLoad.Enabled = True
        mnuSave.Enabled = True
    Else
        mnuLoad.Enabled = False
        mnuSave.Enabled = False
    End If
    
    'Set random or predefined
    If blnRandomTurn = True Then RandomTurn
    
    'Play Game start sound
    PlayWave Play, &H1
    
    'Tell player who starts first
    'MsgBox Turn & " starts game", vbInformation, App.Title

    SetPlayerName

    SetAIBoolean

    If GameMode = "A" Then 'One Player
        Moves = 0 'Initialize Moves
        Delay = 0 'Wait AI click button
    
        'Set AI thinking time
        ThinkTime = ((Rnd() * 100) Mod 5) + 2 'at least 2 sec
    
        If Turn = PlayerColour Then
            EnableButton
            WaitPlayer
        Else
            'AI Start Moves
            DelayTimer.Enabled = True
        End If
    ElseIf GameMode = "B" Then
        EnableButton
        WaitPlayer
    Else 'GameMode= "S"
        'Generate AI
        AINum(BLUE) = GenerateAI
        AINum(GREEN) = GenerateAI
        AINum(RED) = GenerateAI
        AINum(YELLOW) = GenerateAI
        
        Moves = 0 'Initialize Moves
        Delay = 0 'Wait AI click button
    
        'Set AI thinking time
        ThinkTime = 2 ' Wait 2 seconds
        
        'For iCol = 1 To 4
        '    Win(iCol) = 0
        '    blnVictory(iCol) = False
        'Next
    
        'AI Start Moves
        DelayTimer.Enabled = True
    End If

    If cmdRoll(Turn).Enabled = True Then
        cmdRoll(Turn).SetFocus
    End If
    
    ' For Debug
    If Text1.Visible = True Then
        Text1.Enabled = True
        Command1.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting
    'MsgBox "Thank you for testing" & Chr(13) & "Please send feedback to" & Chr(13) & "yiphoon@yahoo.com", vbInformation
    Unload frmChoose
    Unload frmRules
    Unload frmSettings
    'End
End Sub

Private Sub mnuExit_Click()
    SaveSetting
    'MsgBox "Thank you for testing" & Chr(13) & "Please send feedback to" & Chr(13) & "yiphoon@yahoo.com", vbInformation
    End 'Unload Me
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Programmed by Aeric Poon" & Chr(13) & "Last updated on " & App.ProductName & Chr(13) & "Use 800X600 screen size", vbInformation, App.Title
End Sub

Private Sub mnuDebug_Click()
    Text1.Visible = True
    Command1.Visible = True
    mnuDebug.Enabled = False
End Sub

Private Sub mnuMenu_Click()
'    If mnuMenu.Checked = True Then
'        mnuMenu.Checked = False
'        mnuFile.Visible = True
'        mnuGame.Visible = True
'        mnuOption.Visible = True
'        mnuHelp.Visible = True
'        ShowMenu = True
'        If Me.WindowState = 0 Then Me.Height = Me.Height + 50
'    Else
        mnuMenu.Checked = True
        mnuFile.Visible = False
        mnuGame.Visible = False
        mnuOption.Visible = False
        mnuHelp.Visible = False
        ShowMenu = False
        If Me.WindowState = 0 Then Me.Height = Me.Height - 300
'    End If
End Sub

Private Sub mnuRules_Click()
    frmRules.Show vbModal
End Sub

Private Sub picBoard_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If ShowMenu = True Then
            mnuFile.Visible = False
            mnuGame.Visible = False
            mnuOption.Visible = False
            mnuHelp.Visible = False
            mnuMenu.Checked = True
            ShowMenu = False
            'BlinkTimer.Enabled = Not BlinkTimer.Enabled
            If cmdRoll(Turn).Enabled = True Then
                cmdRoll(Turn).SetFocus
            End If
            If Me.WindowState = 0 Then Me.Height = Me.Height - 300
        Else
            mnuFile.Visible = True
            mnuGame.Visible = True
            mnuOption.Visible = True
            mnuHelp.Visible = True
            mnuMenu.Checked = False
            ShowMenu = True
            If Me.WindowState = 0 Then Me.Height = Me.Height + 300
        End If
    End If
End Sub

Private Function Direction()
    Dim FlyPlane As Integer
    Dim PlaneNo As Integer
    Dim D1 As Integer
    Dim D2 As Integer
    Dim GoWin As Integer

On Error GoTo err
    FlyPlane = COLOUR(Turn, Go)
    PlaneNo = Turn * 10

    Select Case Turn
        Case BLUE
            D1 = 58
            D2 = 66
            GoWin = 4
        Case GREEN
            D1 = 19
            D2 = 71
            GoWin = 1
        Case RED
            D1 = 32
            D2 = 76
            GoWin = 2
        Case YELLOW
            D1 = 45
            D2 = 81
            GoWin = 3
    End Select

    If FlyPlane = D1 Or FlyPlane = D2 Then
        IMG(Turn, Go) = GoWin
        Plane(PlaneNo + Go).Picture = picPlane(PlaneNo + GoWin).Picture
    Else
        If blnReverse(Turn, Go) = False Then
            IMG(Turn, Go) = PathDirection(FlyPlane)
            Plane(PlaneNo + Go).Picture = picPlane(PlaneNo + IMG(Turn, Go)).Picture
        Else
            IMG(Turn, Go) = GoWin + 2
            If IMG(Turn, Go) > 4 Then IMG(Turn, Go) = IMG(Turn, Go) - 4
            Plane(PlaneNo + Go).Picture = picPlane(PlaneNo + IMG(Turn, Go)).Picture
        End If
    End If

Exit Function
err:
    MsgBox Error & Chr(13) & "AI (" & Player(Turn) & ") Case " & AINum(Turn), vbExclamation, "Direction"
End Function

Private Function PathDirection(Position As Integer) As Integer
Select Case Position
Case 1, 5, 22, 23, 24, 25, 29, 30, 31, 32, 33, 34, 39, 40, 41, 66, 67, 68, 69, 70
    PathDirection = 1 'South
Case 2, 6, 35, 36, 37, 38, 42, 43, 44, 45, 46, 47, 52, 53, 54, 71, 72, 73, 74, 75
    PathDirection = 2 'West
Case 3, 7, 13, 14, 15, 48, 49, 50, 51, 55, 56, 57, 58, 59, 60, 76, 77, 78, 79, 80
    PathDirection = 3 'North
Case 4, 8, 9, 10, 11, 12, 16, 17, 18, 19, 20, 21, 26, 27, 28, 61, 62, 63, 64, 65 '61 to 65
    PathDirection = 4 'East
End Select
End Function

Private Sub CheckKick()
    Dim PlaneNum As Integer
    Dim i As Integer

    For i = BLUE To YELLOW
        If Turn <> i Then
            PlaneNum = PlaneNum + CountPlane(i, StopPath) 'Skip same colour
            If PlaneNum > 1 Then
                Kick Turn 'Kick by Occupied planes
                Exit Sub
            End If
            If PlaneNum = 1 Then
                Kick i 'Kick Occupied plane
                Exit Sub
            End If
        End If
    Next
End Sub

Private Function CountPlane(ChkColour As Integer, Position As Integer) As Integer
    Dim k As Integer

    For k = 1 To 4 'Occupied by how many blue plane
        If Position = COLOUR(ChkColour, k) Then
            CountPlane = CountPlane + 1
        End If
    Next
End Function

Private Sub Kick(pintColour As Integer)
    Dim PlaneNum As Integer 'Occupied Plane
    Dim PlaneColour As Integer
    Dim n As Integer

    PlayWave OhOu, &H1 'Play sound
    
    PlaneColour = pintColour * 10
    'Check which plane occupied
    For n = 1 To 4
        If StopPath = COLOUR(pintColour, n) Then
            PlaneNum = n
            Exit For
        End If
    Next
    
    Plane(PlaneColour + PlaneNum).Picture = picPlane(PlaneColour + pintColour).Picture
    Plane(PlaneColour + PlaneNum).Left = Base(PlaneColour + PlaneNum).Left
    Plane(PlaneColour + PlaneNum).Top = Base(PlaneColour + PlaneNum).Top
    IMG(pintColour, PlaneNum) = pintColour
    COLOUR(pintColour, PlaneNum) = pintColour 'Back to starting point
End Sub

Private Sub NextTurn()
    Moves = 0 'Initialize Moves
    Delay = 0 'Wait AI click button
    
    intRepeatRoll = 0
    
    If GameMode = "S" Then
        ThinkTime = 2 'Wait 2 seconds
    Else
        ThinkTime = ((Rnd() * 100) Mod 5) + 2 'Set AI wasting time (Min=2, Max 6)
    End If
    
    If blnVictory(BLUE) And _
        blnVictory(GREEN) And _
        blnVictory(RED) And _
        blnVictory(YELLOW) Then
        If vbYes = MsgBox("Restart Game?", vbInformation + vbYesNo, "All Missions accomplished!") Then
            MsgBox "The Game should Reset!", vbOKOnly, App.Title
            MsgBox "Thank you for Testing." + Chr(13) + "Please send feedback to yiphoon@yahoo.com", vbInformation, App.Title
            mnuEnd_Click
        'Else
            
        End If
        Exit Sub
    End If

    Turn = Turn + 1 ' Set Turn to Next COLOUR
    If Turn = 5 Then Turn = 1
    
    If blnVictory(Turn) Then
        ' If during this COLOUR's turn, this COLOUR has won
        ' then next COLOUR's turn
        NextTurn 'No repeat
        Exit Sub
    Else
        ' Else Enable button if not AI
        cmdRoll(Turn).Caption = "Roll"
        If blnAI(Turn) = True Then
            DelayTimer.Enabled = True
        Else
            cmdRoll(Turn).Enabled = True
            cmdRoll(Turn).SetFocus
            WaitPlayer
        End If
    End If

'Debug Purpose
'    Text1.Enabled = True
'    Command1.Enabled = True
End Sub

Private Sub SameTurn()
    Moves = 0 'Initialize Moves
    Delay = 0 'Wait AI click button

    intRepeatRoll = intRepeatRoll + 1
    
    If GameMode = "S" Then
        ThinkTime = 2 'Wait 2 seconds
    Else
        ThinkTime = ((Rnd() * 100) Mod 5) + 2 'Set AI wasting time (Min=2, Max 6)
    End If
    
    If blnVictory(BLUE) And _
        blnVictory(GREEN) And _
        blnVictory(RED) And _
        blnVictory(YELLOW) Then
        If vbYes = MsgBox("Restart Game?", vbInformation + vbYesNo, "All Missions accomplished!") Then
            MsgBox "The Game should Reset!", vbOKOnly, App.Title
            MsgBox "Thank you for Testing." + Chr(13) + "Please send feedback to yiphoon@yahoo.com", vbInformation, App.Title
            mnuEnd_Click
        'Else
            
        End If
        Exit Sub
    End If

    'Turn = Turn + 1 ' Set Turn to Next COLOUR
    'If Turn = 5 Then Turn = 1
    
    If blnVictory(Turn) Then
        ' If during this COLOUR's turn, this COLOUR has won
        ' then next COLOUR's turn
        NextTurn 'No repeat
        Exit Sub
    Else
        ' Else Enable button if not AI
        cmdRoll(Turn).Caption = "Roll"
        If blnAI(Turn) = True Then
            DelayTimer.Enabled = True
        Else
            cmdRoll(Turn).Enabled = True
            cmdRoll(Turn).SetFocus
            WaitPlayer
        End If
    End If

'Debug Purpose
'    Text1.Enabled = True
'    Command1.Enabled = True
End Sub

Private Sub ShortCutTimer_Timer()
    Dim Destination As Integer
    
    Select Case Turn
        Case BLUE
            intFlying = intFlying + 1
        Case GREEN
            intFlying = intFlying - 1
        Case RED
            intFlying = intFlying - 1
        Case YELLOW
            intFlying = intFlying + 1
    End Select
    
    ShortCut
    If CheckEndShortcut() = True Then
        ShortCutTimer.Enabled = False
        Select Case Turn
            Case BLUE
                Destination = 38
            Case GREEN
                Destination = 51
            Case RED
                Destination = 12
            Case YELLOW
                Destination = 25
        End Select
        COLOUR(Turn, Go) = Destination
        StopPath = Destination 'COLOUR(Turn, Go)
        If blnKickPlane = True Then
            CheckKick
        End If
        Direction              ' Dont want turn ?
        'NextTurn
        If Steps = 6 And blnRollAgain = True Then  'Roll again if six
            If intRepeatRoll < intMaxRoll Then
                SameTurn
                'intRepeatRoll = intRepeatRoll + 1
            Else
                NextTurn ' No repeat
            End If
        Else
            NextTurn ' No repeat
        End If
    End If
End Sub

Private Sub ShortCut()
On Error GoTo err
    'Change Direction
    Select Case Turn
        Case BLUE
            For intFlying = intFlying To intFlying + 10 Step 5
                Plane(BLUE * 10 + Go).Move Path(26).Left, intFlying
                'Plane(blue*10+Go).Refresh
            Next
        Case GREEN
            For intFlying = intFlying To intFlying - 10 Step -5
                Plane(GREEN * 10 + Go).Move intFlying, Path(39).Top
                'Plane(green*10+Go).Refresh
            Next
        Case RED
            For intFlying = intFlying To intFlying - 10 Step -5
                Plane(RED * 10 + Go).Move Path(52).Left, intFlying
                'Plane(red*10+Go).Refresh
            Next
        Case YELLOW
            For intFlying = intFlying To intFlying + 10 Step 5
                Plane(YELLOW * 10 + Go).Move intFlying, Path(13).Top
                'Plane(yellow*10+Go).Refresh
            Next
    End Select
    Exit Sub
err:
    MsgBox Error & Chr(13) & gstrObject & " (" & Player(Turn) & ") = " & Go, vbExclamation, "Shortcut"
End Sub

Private Sub DelayTimer_Timer() 'Wait AI to Roll die (wasting time)
    If Delay = 0 Then
        ' lblMessage.Caption = ""
        fraMessage.Visible = False
    End If
    If Delay = 1 Then
        WaitPlayer
    End If
    If Delay < ThinkTime Then
        Delay = Delay + 1
    Else
        DelayTimer.Enabled = False
        cmdRoll_Click (Turn)
    End If
End Sub

Private Sub AIThinkTimer_Timer() 'Delay a while before click a selected plane
    Select Case Turn
        Case BLUE
            fraMessage.BackColor = vbBlue
        Case GREEN
            fraMessage.BackColor = vbGreen
        Case RED
            fraMessage.BackColor = vbRed
        Case YELLOW
            fraMessage.BackColor = vbYellow
    End Select
    
    If Delay = 0 Then 'Wait a second before show message
        'lblMessage.Caption = ""
        fraMessage.Visible = False
    End If
    If Delay = 1 Then 'Wait a second before show message
        lblMessage.Caption = Player(Turn) & " (AI) is thinking..."
        fraMessage.Visible = True
    End If
    If Delay < ThinkTime Then
        Delay = Delay + 1
    Else
        AIThinkTimer.Enabled = False 'Mandatory
        'lblMessage.Caption = ""
        fraMessage.Visible = False
        ' AI Clicks
        blnClick = True
    End If
End Sub

Private Function EnablePlane()
Dim i As Integer
On Error GoTo err
    For i = 1 To 4
        'CheckMove
        If blnColour(Turn, i) = True Then
            Plane(Turn * 10 + i).Enabled = True
        End If
    Next
Exit Function
err:
    MsgBox Error & Chr(13) & Player(Turn) & " (Go = " & Go & ")", vbExclamation, "Enable" & gstrObject
End Function

Private Function DisablePlane()
Dim i As Integer
On Error GoTo err
    For i = 1 To 4
        Plane(Turn * 10 + i).Enabled = False
    Next
Exit Function
err:
    MsgBox Error & Chr(13) & Player(Turn) & " (Go = " & Go & ")", vbExclamation, "Disable" & gstrObject
End Function

Private Sub WaitPlayer()
    Select Case Turn
        Case BLUE
            fraMessage.BackColor = vbBlue
        Case GREEN
            fraMessage.BackColor = vbGreen
        Case RED
            fraMessage.BackColor = vbRed
        Case YELLOW
            fraMessage.BackColor = vbYellow
    End Select
    
    If blnAI(Turn) Then
        'Show Message that waiting for AI to Roll die
        lblMessage.Caption = "Waiting " & Player(Turn) & " (AI)..."
    Else
        'Wait for Player to Roll die
        lblMessage.Caption = Player(Turn) & "'s Turn..."
    End If
    fraMessage.Visible = True
End Sub

Public Function CheckAnyJump() As Integer
    Dim JumpNo As Integer
    Dim j As Integer

    Select Case Turn
        Case BLUE
            JumpNo = 2
        Case GREEN
            JumpNo = 3
        Case RED
            JumpNo = 0
        Case YELLOW
            JumpNo = 1
    End Select

    For j = 1 To 4
        If COLOUR(Turn, j) > 4 Then ' Check if plane enabled?
            If (COLOUR(Turn, j) + Steps) Mod 4 = JumpNo Then
                CheckAnyJump = j
                Exit Function
            End If
        End If
    Next
End Function

Public Function CheckAnyKick() As Integer
    Dim KickNo As Integer
    Dim k As Integer 'k=1 is Plane 1 of COLOUR
    Dim i As Integer 'i=4 is Plane 4 of Other COLOUR
    Dim l As Integer 'l=2 is GREEN for Looping COLOUR
    
    
    Select Case Turn
        Case BLUE
            KickNo = 2
        Case GREEN
            KickNo = 3
        Case RED
            KickNo = 0
        Case YELLOW
            KickNo = 1
    End Select
    
    For k = 1 To 4
        ' Is Plane enabled?
        ' Used blnColour(Turn, k)?
        ' or IsBlinkingPlane?
        If COLOUR(Turn, k) > 4 Then
            For l = 1 To 4
                If l <> Turn Then ' Skip same COLOUR
                    For i = 1 To 4
                        If COLOUR(Turn, k) + Steps = COLOUR(l, i) Then
                            'If it will jump then do not choose that plane
                            If Not (COLOUR(Turn, k) + Steps) Mod 4 = KickNo Then
                                CheckAnyKick = k
                                Exit Function
                            End If
                        End If
                    Next
                End If
            Next
        End If
    Next
End Function

Public Function CheckAnyShortcut() As Integer
    Dim s As Integer
    Dim ShortCutNo As Integer
    
    ShortCutNo = ((Turn Mod 4) + 1) * 13
    For s = 1 To 4
        If COLOUR(Turn, s) > 4 Then ' Is Plane enabled?
            If COLOUR(Turn, s) + Steps = ShortCutNo Then
                CheckAnyShortcut = s
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsBlinkingPlane(TurnNum As Integer) As Boolean
    'Every plane ready to go if not goal
    'If no checking of whether COLOUR(1,1) = 0 then Goaled plane still can move
    'This will cause Plane moves to Path(1) which is not exist (Goaled Plane = 0)
    'If COLOUR(1,1) > 4 Then
    If COLOUR(Turn, TurnNum) < 5 Then
        If COLOUR(Turn, TurnNum) = 0 Then
            IsBlinkingPlane = False
        Else
            If Steps = 6 Then
                IsBlinkingPlane = True
            Else '==========DieNumber Not 6==========
                If blnFirstRound = True And blnSixOnlyStart = False Then
                    IsBlinkingPlane = True
                Else
                    IsBlinkingPlane = False
                End If
            End If
        End If
    Else
        IsBlinkingPlane = True
    End If
End Function

Public Function IsInsideBase(PlaneNum As Integer) As Boolean
    If COLOUR(Turn, PlaneNum) = Turn Then IsInsideBase = True
End Function

Public Function IsGoal(PlaneNum As Integer) As Boolean
    If COLOUR(Turn, PlaneNum) = 0 Then IsGoal = True
End Function

Public Function Smallest() As Integer
    'This function return the plane that left behind (Not inside Base)
    'Note: remember to check plane Enabled=False e.g COLOUR(1,1) = 1
    Dim Color(1 To 4) As Integer
    Dim Ea As Integer
    Dim Eb As Integer
    Dim i As Integer
    
    For i = 1 To 4
        Color(i) = COLOUR(Turn, i)
    Next

    Ea = 0
    Eb = 0
    
    If Color(1) > 4 Then 'E1
        If Color(2) > 4 Then 'E1E2
            If Color(1) <= Color(2) Then 'E1<=E2
                Ea = 1 'Blue(1)
            Else
                Ea = 2 'Blue(2)
            End If
        Else 'E1D2
            Ea = 1 'Blue(1)
        End If
    Else 'D1
        If Color(2) > 4 Then 'D1E2
            Ea = 2 'Blue(2)
        'Else 'D1D2
        '    Ea = 0
        End If
    End If
    If Color(3) > 4 Then 'E3
        If Color(4) > 4 Then 'E3E4
            If Color(3) <= Color(4) Then 'E3<=E4
                Eb = 3 'Blue(3)
            Else
                Eb = 4 'Blue(4)
            End If
        Else 'E3D4
            Eb = 3 'Blue(3)
        End If
    Else 'D3
        If Color(4) > 4 Then 'D3E4
            Eb = 4 'Blue(4)
        'Else 'D3D4
        '    Eb = 0
        End If
    End If
        
    If Ea > 0 Then 'Ea
        If Eb > 0 Then 'EaEb
            If Color(Ea) <= Color(Eb) Then 'Ea<=Eb
                Smallest = Ea 'Yellow(a)
            Else
                Smallest = Eb 'Yellow(b)
            End If
        Else 'EaDb
            Smallest = Ea 'Yellow(a)
        End If
    Else 'Da
        If Eb > 0 Then 'DaEb
            Smallest = Eb 'Yellow(b)
        'Else 'DaDb
        '    Smallest = 0
        End If
    End If
'    If Smallest = 0 Then
'        'No plane is moveable/ all planes victory
'    End If
    If blnFirstRound = True And blnSixOnlyStart = False Then
        Smallest = 1
    End If
End Function

Public Function Largest() As Integer
    'This function return the plane that is most front
    'Note: remember to check plane Enabled=False e.g COLOUR(1,1) = 1
    Dim Color(1 To 4) As Integer
    Dim Ea As Integer
    Dim Eb As Integer
    Dim i As Integer
    
    For i = 1 To 4
        Color(i) = COLOUR(Turn, i)
    Next
    
    Ea = 0
    Eb = 0
    
    If Color(1) > 4 Then 'E1
        If Color(2) > 4 Then 'E1E2
            If Color(1) >= Color(2) Then 'E1>=E2
                Ea = 1 'Blue(1)
            Else
                Ea = 2 'Blue(2)
            End If
        Else 'E1D2
            Ea = 1 'Blue(1)
        End If
    Else 'D1
        If Color(2) > 4 Then 'D1E2
            Ea = 2 'Blue(2)
        'Else 'D1D2
        '    Ea = 0
        End If
    End If
    If Color(3) > 4 Then 'E3
        If Color(4) > 4 Then 'E3E4
            If Color(3) >= Color(4) Then 'E3>=E4
                Eb = 3 'Blue(3)
            Else
                Eb = 4 'Blue(4)
            End If
        Else 'E3D4
            Eb = 3 'Blue(3)
        End If
    Else 'D3
        If Color(4) > 4 Then 'D3E4
            Eb = 4 'Blue(4)
        'Else 'D3D4
        '    Eb = 0
        End If
    End If
        
    If Ea > 0 Then 'Ea
        If Eb > 0 Then 'EaEb
            If Color(Ea) >= Color(Eb) Then 'Ea>=Eb
                Largest = Ea 'Blue(a)
            Else
                Largest = Eb 'Blue(b)
            End If
        Else 'EaDb
            Largest = Ea 'Blue(a)
        End If
    Else 'Da
        If Eb > 0 Then 'DaEb
            Largest = Eb 'Blue(b)
        'Else 'DaDb
        '    Largest = 0
        End If
    End If
'    If Largest = 0 Then
'        'No plane is moveable/ all planes victory
'    End If
    If blnFirstRound = True And blnSixOnlyStart = False Then
        Largest = 1
    End If
End Function

Private Sub EnableButton()
    cmdRoll(Turn).Enabled = True
End Sub

Private Function CheckEndShortcut() As Boolean
    Select Case Turn
    Case BLUE
        If Plane(BLUE * 10 + Go).Top >= Path(38).Top Then CheckEndShortcut = True
    Case GREEN
        If Plane(GREEN * 10 + Go).Left <= Path(51).Left Then CheckEndShortcut = True
    Case RED
        If Plane(RED * 10 + Go).Top <= Path(12).Top Then CheckEndShortcut = True
    Case YELLOW
        If Plane(YELLOW * 10 + Go).Left >= Path(25).Left Then CheckEndShortcut = True
    End Select
End Function

Private Sub LoadRule()
Dim strRule1 As String
Dim strRule2 As String
Dim strRule3 As String
Dim strRule4 As String

On Error GoTo err

ReadText "Set", 8, strRule1
ReadText "Set", 9, strRule2
ReadText "Set", 10, strRule3
ReadText "Set", 11, strRule4

If strRule1 = "0" Then blnKickPlane = False Else blnKickPlane = True
If strRule2 = "0" Then blnEnableJump = False Else blnEnableJump = True
If strRule3 = "0" Then blnSixOnlyStart = False Else blnSixOnlyStart = True
If strRule4 = "0" Then blnRollAgain = False Else blnRollAgain = True
    
Exit Sub
err:
    AppendText "Error", Now & " " & "(LoadRule) Error: " & Error
End Sub

Private Sub LoadSetting()
Dim strNext As String
Dim strPlayer As String
Dim strName As String
Dim strStart As String

On Error GoTo err

ReadText "Set", 1, gstrObject
ReadText "Set", 2, GameMode
ReadText "Set", 3, strNext      ' Next Turn
ReadText "Set", 4, strPlayer    ' Default Player Colour
ReadText "Set", 5, strName      ' Turn use Colour or Custom as Name
ReadText "Set", 6, strStart     ' Game Start Colour
ReadText "Set", 7, PlayerName

If gstrObject = "" Then gstrObject = "Plane"

If GameMode = "B" Then
    mnuOne.Checked = False
    mnuFour.Checked = True
    mnuSim.Checked = False
    GameMode = "B"
    DelayTimer.Interval = 1000
    FlyTimer.Interval = 1000
    AIThinkTimer.Interval = 1000
ElseIf GameMode = "S" Then
    mnuOne.Checked = False
    mnuFour.Checked = False
    mnuSim.Checked = True
    GameMode = "S"
    DelayTimer.Interval = 500
    FlyTimer.Interval = 500
    AIThinkTimer.Interval = 500
Else
    mnuOne.Checked = True
    mnuFour.Checked = False
    mnuSim.Checked = False
    'frmChoose.Show vbModal
    GameMode = "A"
    DelayTimer.Interval = 1000
    FlyTimer.Interval = 1000
    AIThinkTimer.Interval = 1000
End If

Select Case UCase(strNext)
Case "BLUE"
    Turn = BLUE
Case "GREEN"
    Turn = GREEN
Case "RED"
    Turn = RED
Case "YELLOW"
    Turn = YELLOW
Case Else
    Turn = gstrDefault
End Select
blnRandomTurn = False

Select Case UCase(strPlayer)
Case "BLUE"
    gstrDefault = BLUE
Case "GREEN"
    gstrDefault = GREEN
Case "RED"
    gstrDefault = RED
Case "YELLOW"
    gstrDefault = YELLOW
Case Else
    gstrDefault = BLUE
End Select
PlayerColour = gstrDefault

If CInt(strName) = 0 Then
    gblnUseName = False
Else
    gblnUseName = True
End If

If CInt(strStart) > 0 And CInt(strStart) < 5 Then
    gintStartColour = CInt(strStart)
Else
    gintStartColour = 0
End If
    
Exit Sub
err:
    AppendText "Error", Now & " " & "(LoadSetting) Error: " & Error
End Sub

Private Sub LoadGame()
LoadSetting
LoadRule
LoadPlayerInfo
SetAIBoolean
'InitPlane
'Default Game Mode
' This is not First round
intStartClick = 0
'blnFirstRound = True
Exit Sub
err:
    AppendText "Error", Now & " " & "(LoadGame) Error: " & Error
End Sub

Private Sub SaveGame()
Dim strNext As String
Dim strPlayer As String

On Error GoTo err

Open App.Path & "\Set.txt" For Output As #4
    If gstrObject = "" Then gstrObject = "Plane"
    If GameMode = "" Then GameMode = "A"
    Select Case Turn
        Case BLUE
            strNext = "BLUE"
        Case GREEN
            strNext = "GREEN"
        Case RED
            strNext = "RED"
        Case YELLOW
            strNext = "YELLOW"
        Case Else
            strNext = "BLUE"
    End Select
    Select Case PlayerColour
        Case BLUE
            strPlayer = "BLUE"
        Case GREEN
            strPlayer = "GREEN"
        Case RED
            strPlayer = "RED"
        Case YELLOW
            strPlayer = "YELLOW"
        Case Else
            strPlayer = "BLUE"
    End Select
    
    Write #4, "[Set]"
    Write #4, gstrObject
    Write #4, GameMode
    Write #4, strNext
    Write #4, strPlayer
    If gblnUseName = False Then
        Write #4, 0
    Else
        Write #4, 1
    End If
    If gintStartColour > 0 And gintStartColour < 5 Then
        Write #4, gintStartColour
    Else
        Write #4, 0
    End If
    
    Write #4, PlayerName
    'Game Rules
    If blnKickPlane = False Then Write #4, 0 Else Write #4, 1
    If blnEnableJump = False Then Write #4, 0 Else Write #4, 1
    If blnSixOnlyStart = False Then Write #4, 0 Else Write #4, 1
    If blnRollAgain = False Then Write #4, 0 Else Write #4, 1
    Close

Open App.Path & "\Saved.txt" For Output As #5
    Write #5, Trim$(frmSettings.txtTurn(BLUE).Text), COLOUR(BLUE, 1), COLOUR(BLUE, 2), COLOUR(BLUE, 3), COLOUR(BLUE, 4), AINum(BLUE)
    Write #5, Trim$(frmSettings.txtTurn(GREEN).Text), COLOUR(GREEN, 1), COLOUR(GREEN, 2), COLOUR(GREEN, 3), COLOUR(GREEN, 4), AINum(GREEN)
    Write #5, Trim$(frmSettings.txtTurn(RED).Text), COLOUR(RED, 1), COLOUR(RED, 2), COLOUR(RED, 3), COLOUR(RED, 4), AINum(RED)
    Write #5, Trim$(frmSettings.txtTurn(YELLOW).Text), COLOUR(YELLOW, 1), COLOUR(YELLOW, 2), COLOUR(YELLOW, 3), COLOUR(YELLOW, 4), AINum(YELLOW)
Close #5

Exit Sub
err:
    AppendText "Error", Now & " " & "(SaveGame) Error: " & Error
End Sub

Private Sub SaveSetting()
Dim strNext As String
Dim strPlayer As String

On Error GoTo err

Open App.Path & "\Set.txt" For Output As #4
    If gstrObject = "" Then gstrObject = "Plane"
    If GameMode = "" Then GameMode = "A"
    Select Case Turn
        Case BLUE
            strNext = "BLUE"
        Case GREEN
            strNext = "GREEN"
        Case RED
            strNext = "RED"
        Case YELLOW
            strNext = "YELLOW"
        Case Else
            strNext = "BLUE"
    End Select
    Select Case PlayerColour
        Case BLUE
            strPlayer = "BLUE"
        Case GREEN
            strPlayer = "GREEN"
        Case RED
            strPlayer = "RED"
        Case YELLOW
            strPlayer = "YELLOW"
        Case Else
            strPlayer = "BLUE"
    End Select
    
    Write #4, "[Set]"
    Write #4, gstrObject
    Write #4, GameMode
    Write #4, strNext
    Write #4, strPlayer
    If gblnUseName = False Then
        Write #4, 0
    Else
        Write #4, 1
    End If
    If gintStartColour > 0 And gintStartColour < 5 Then
        Write #4, gintStartColour
    Else
        Write #4, 0
    End If
    
    Write #4, PlayerName
    'Game Rules
    If blnKickPlane = False Then Write #4, 0 Else Write #4, 1
    If blnEnableJump = False Then Write #4, 0 Else Write #4, 1
    If blnSixOnlyStart = False Then Write #4, 0 Else Write #4, 1
    If blnRollAgain = False Then Write #4, 0 Else Write #4, 1
    Close

Exit Sub
err:
    AppendText "Error", Now & " " & "(SaveSetting) Error: " & Error
End Sub

Private Sub SetAIBoolean()
    If UCase(GameMode) = "S" Then
        GameMode = "S"
        blnAI(BLUE) = True
        blnAI(GREEN) = True
        blnAI(RED) = True
        blnAI(YELLOW) = True
    ElseIf UCase(GameMode) = "B" Then
        GameMode = "B"
        blnAI(BLUE) = False
        blnAI(GREEN) = False
        blnAI(RED) = False
        blnAI(YELLOW) = False
    Else
        ' UCase(GameMode) = "A" then
        GameMode = "A"
        blnAI(BLUE) = True
        blnAI(GREEN) = True
        blnAI(RED) = True
        blnAI(YELLOW) = True
        blnAI(PlayerColour) = False
    End If
End Sub

Private Sub LoadPlayerInfo()
    Open App.Path & "\Saved.txt" For Input As #3
        If Not EOF(3) Then
            Input #3, Player(BLUE), COLOUR(BLUE, 1), COLOUR(BLUE, 2), COLOUR(BLUE, 3), COLOUR(BLUE, 4), AINum(BLUE)
            Input #3, Player(GREEN), COLOUR(GREEN, 1), COLOUR(GREEN, 2), COLOUR(GREEN, 3), COLOUR(GREEN, 4), AINum(GREEN)
            Input #3, Player(RED), COLOUR(RED, 1), COLOUR(RED, 2), COLOUR(RED, 3), COLOUR(RED, 4), AINum(RED)
            Input #3, Player(YELLOW), COLOUR(YELLOW, 1), COLOUR(YELLOW, 2), COLOUR(YELLOW, 3), COLOUR(YELLOW, 4), AINum(YELLOW)
        End If
    Close #3
End Sub

Private Sub SetPlanePosition()
Dim iCol As Integer
Dim i As Integer
Dim PathNo As Integer
                
For iCol = 1 To 4 'BLUE To YELLOW
    Win(iCol) = 0
    blnVictory(iCol) = False
    For i = 1 To 4
        PathNo = COLOUR(iCol, i)
    
        ' Eg. For i = 1
        ' IMG(BLUE, 1) = PathDirection(1)   -> means Blue1 = South
        ' Plane(11).Picture = picPlane(1, IMG(BLUE, 1)).Picture
        ' blnAI(i) = True

        If PathNo = 0 Then
            Win(iCol) = Win(iCol) + 1
            If iCol - 1 = 0 Then
                IMG(iCol, i) = 4
            Else
                IMG(iCol, i) = iCol - 1
            End If
            Plane(iCol * 10 + i).Picture = picPlane(iCol * 10 + IMG(iCol, i)).Picture
            Plane(iCol * 10 + i).Left = Goal(iCol).Left
            Plane(iCol * 10 + i).Top = Goal(iCol).Top
            
            If Win(iCol) = 4 Then
                'no more turn
                blnVictory(iCol) = True
            End If
        Else
            IMG(iCol, i) = PathDirection(PathNo)
            If PathNo > 4 Then
                Plane(iCol * 10 + i).Picture = picPlane(iCol * 10 + IMG(iCol, i)).Picture
            Else
                Plane(iCol * 10 + i).Picture = picPlane(iCol * 10 + iCol).Picture
            End If
            
            If PathNo > 4 Then
                Plane(iCol * 10 + i).Left = Path(PathNo).Left
                Plane(iCol * 10 + i).Top = Path(PathNo).Top
            End If
        End If

        ' All Planes are not blinking
        blnColour(iCol, i) = False
        ' Direction
    Next
Next

End Sub

Private Sub SetPlayerName()
    Dim i As Integer
    
    If frmSettings.optName(1).Value = True Then
        For i = 1 To 4
            Player(i) = frmSettings.txtTurn(i).Text
        Next
    Else
        Player(BLUE) = "BLUE"
        Player(GREEN) = "GREEN"
        Player(RED) = "RED"
        Player(YELLOW) = "YELLOW"
    End If

    If Trim(frmSettings.txtPlayerName.Text) <> "" And GameMode = "A" Then
        PlayerName = Trim(frmSettings.txtPlayerName.Text)
        Player(PlayerColour) = PlayerName
    End If
End Sub
