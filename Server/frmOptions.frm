VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Options Window"
   ClientHeight    =   6870
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog comDiag 
      Left            =   4200
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.maz"
      DialogTitle     =   "Save/Open Map"
      Filter          =   "WOTA Maze Files (*.maz) |*.maz|"
      InitDir         =   "app.path"
   End
   Begin VB.Frame framQuest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quest Only Options"
      Height          =   3015
      Left            =   8880
      TabIndex        =   6
      Top             =   3720
      Width           =   4095
      Begin VB.CheckBox chkMove 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Can Move a Wooden Pillar On"
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkJump 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Can Jump To a Wooden Pillar"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkGoTo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable Go To"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Text            =   "Level Name"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtGoto 
         Height          =   285
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.CheckBox chkRandom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Random Battle"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkBoss 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Boss Battle (Place Boss On Tile)"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblX 
         BackStyle       =   0  'Transparent
         Caption         =   "Goto Tile:"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame framMaze 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maze Tiles and Maze Only Options"
      Height          =   3975
      Left            =   7800
      TabIndex        =   1
      Top             =   0
      Width           =   5175
      Begin VB.CheckBox chkEndTile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "This Tile Ends Maze (Place On All Finishing Tiles)"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   4215
      End
      Begin VB.TextBox txtMazeStart 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Tile (Maze Only):"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   1770
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   18
         Left            =   3000
         Picture         =   "frmOptions.frx":0000
         Top             =   2640
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   19
         Left            =   120
         Picture         =   "frmOptions.frx":05B9
         Top             =   2640
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   20
         Left            =   840
         Picture         =   "frmOptions.frx":0B45
         Top             =   2640
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   27
         Left            =   1560
         Picture         =   "frmOptions.frx":10D5
         Top             =   2640
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   28
         Left            =   2280
         Picture         =   "frmOptions.frx":1657
         Top             =   2640
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "frmOptions.frx":1BD2
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   1
         Left            =   840
         Picture         =   "frmOptions.frx":2113
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   2
         Left            =   1560
         Picture         =   "frmOptions.frx":25F3
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   3
         Left            =   2280
         Picture         =   "frmOptions.frx":2B61
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   4
         Left            =   3000
         Picture         =   "frmOptions.frx":30BB
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   5
         Left            =   3720
         Picture         =   "frmOptions.frx":367B
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   6
         Left            =   120
         Picture         =   "frmOptions.frx":3C2E
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   7
         Left            =   840
         Picture         =   "frmOptions.frx":41C8
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   8
         Left            =   1560
         Picture         =   "frmOptions.frx":475C
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   9
         Left            =   2280
         Picture         =   "frmOptions.frx":4CD6
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   10
         Left            =   3000
         Picture         =   "frmOptions.frx":5275
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   11
         Left            =   3720
         Picture         =   "frmOptions.frx":5812
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   12
         Left            =   120
         Picture         =   "frmOptions.frx":5D93
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   13
         Left            =   840
         Picture         =   "frmOptions.frx":630B
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   14
         Left            =   1560
         Picture         =   "frmOptions.frx":6864
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   15
         Left            =   2280
         Picture         =   "frmOptions.frx":6DB9
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   16
         Left            =   3000
         Picture         =   "frmOptions.frx":7336
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   17
         Left            =   3720
         Picture         =   "frmOptions.frx":78A1
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   21
         Left            =   120
         Picture         =   "frmOptions.frx":7E52
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   22
         Left            =   840
         Picture         =   "frmOptions.frx":8297
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   23
         Left            =   1560
         Picture         =   "frmOptions.frx":86B4
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   24
         Left            =   2280
         Picture         =   "frmOptions.frx":8ABC
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   25
         Left            =   3000
         Picture         =   "frmOptions.frx":8EED
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgTile 
         Height          =   375
         Index           =   26
         Left            =   3720
         Picture         =   "frmOptions.frx":9324
         Top             =   2040
         Width           =   375
      End
   End
   Begin VB.Label lblGen 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.2 - 5/23/2003"
      Height          =   195
      Index           =   1
      Left            =   6120
      TabIndex        =   15
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   68
      Left            =   6480
      Picture         =   "frmOptions.frx":976D
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   67
      Left            =   6480
      Picture         =   "frmOptions.frx":9B82
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   750
      Index           =   32
      Left            =   1080
      Picture         =   "frmOptions.frx":9F98
      Top             =   2880
      Width           =   750
   End
   Begin VB.Image imgSprite 
      Height          =   750
      Index           =   31
      Left            =   120
      Picture         =   "frmOptions.frx":A6B2
      Top             =   2760
      Width           =   750
   End
   Begin VB.Image imgSprite 
      Height          =   750
      Index           =   30
      Left            =   4320
      Picture         =   "frmOptions.frx":ADD7
      Top             =   2400
      Width           =   750
   End
   Begin VB.Image imgSprite 
      Height          =   375
      Index           =   29
      Left            =   3840
      Picture         =   "frmOptions.frx":B47F
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   375
      Index           =   28
      Left            =   3360
      Picture         =   "frmOptions.frx":B8AE
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   375
      Index           =   27
      Left            =   2880
      Picture         =   "frmOptions.frx":BD4B
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   750
      Index           =   26
      Left            =   2040
      Picture         =   "frmOptions.frx":C1B6
      Top             =   2400
      Width           =   750
   End
   Begin VB.Image imgSprite 
      Height          =   375
      Index           =   25
      Left            =   1560
      Picture         =   "frmOptions.frx":C934
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   375
      Index           =   24
      Left            =   1080
      Picture         =   "frmOptions.frx":CD20
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   375
      Index           =   23
      Left            =   600
      Picture         =   "frmOptions.frx":D168
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   375
      Index           =   22
      Left            =   120
      Picture         =   "frmOptions.frx":D5A6
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblBosses 
      BackStyle       =   0  'Transparent
      Caption         =   "Bosses (Place on a walkable tile with Boss Enabled)"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image imgSprite 
      Height          =   765
      Index           =   21
      Left            =   600
      Picture         =   "frmOptions.frx":DA55
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   66
      Left            =   2760
      Picture         =   "frmOptions.frx":E083
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   1125
      Index           =   20
      Left            =   120
      Picture         =   "frmOptions.frx":E4A8
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   375
      Index           =   19
      Left            =   480
      Picture         =   "frmOptions.frx":EAF6
      Top             =   120
      Width           =   1125
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   65
      Left            =   6840
      Picture         =   "frmOptions.frx":F0A6
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   64
      Left            =   6120
      Picture         =   "frmOptions.frx":F4E5
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   63
      Left            =   6840
      Picture         =   "frmOptions.frx":F925
      Top             =   960
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   62
      Left            =   6120
      Picture         =   "frmOptions.frx":FD4A
      Top             =   960
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   61
      Left            =   6840
      Picture         =   "frmOptions.frx":1016F
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   60
      Left            =   6120
      Picture         =   "frmOptions.frx":105AA
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   59
      Left            =   3000
      Picture         =   "frmOptions.frx":109EE
      Top             =   960
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   58
      Left            =   4920
      Picture         =   "frmOptions.frx":10E6E
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   57
      Left            =   4920
      Picture         =   "frmOptions.frx":112C6
      Top             =   480
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   56
      Left            =   4560
      Picture         =   "frmOptions.frx":1170B
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   55
      Left            =   4560
      Picture         =   "frmOptions.frx":11B5E
      Top             =   480
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   54
      Left            =   4560
      Picture         =   "frmOptions.frx":11F86
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   1710
      Index           =   18
      Left            =   3960
      Picture         =   "frmOptions.frx":123DF
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   53
      Left            =   5640
      Picture         =   "frmOptions.frx":13071
      Top             =   480
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   52
      Left            =   5640
      Picture         =   "frmOptions.frx":13520
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   51
      Left            =   5280
      Picture         =   "frmOptions.frx":139B5
      Top             =   480
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   50
      Left            =   5280
      Picture         =   "frmOptions.frx":13E67
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   49
      Left            =   3000
      Picture         =   "frmOptions.frx":142F8
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   48
      Left            =   3000
      Picture         =   "frmOptions.frx":14793
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   47
      Left            =   4080
      Picture         =   "frmOptions.frx":14C2B
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   46
      Left            =   4080
      Picture         =   "frmOptions.frx":150BE
      Top             =   480
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   45
      Left            =   4080
      Picture         =   "frmOptions.frx":1556A
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   44
      Left            =   3600
      Picture         =   "frmOptions.frx":15A2B
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   43
      Left            =   3600
      Picture         =   "frmOptions.frx":15EB7
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   42
      Left            =   3600
      Picture         =   "frmOptions.frx":16375
      Top             =   480
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   870
      Index           =   16
      Left            =   2040
      Picture         =   "frmOptions.frx":16822
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   41
      Left            =   3000
      Picture         =   "frmOptions.frx":16E42
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   40
      Left            =   2640
      Picture         =   "frmOptions.frx":17231
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   3195
      Index           =   17
      Left            =   5040
      Picture         =   "frmOptions.frx":17582
      Top             =   4800
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   39
      Left            =   2520
      Picture         =   "frmOptions.frx":1B370
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   38
      Left            =   2520
      Picture         =   "frmOptions.frx":1B7EA
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   37
      Left            =   2040
      Picture         =   "frmOptions.frx":1BC6A
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   36
      Left            =   2040
      Picture         =   "frmOptions.frx":1C0E5
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   35
      Left            =   1560
      Picture         =   "frmOptions.frx":1C570
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   34
      Left            =   1200
      Picture         =   "frmOptions.frx":1CA35
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   33
      Left            =   1560
      Picture         =   "frmOptions.frx":1CF1D
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   32
      Left            =   1200
      Picture         =   "frmOptions.frx":1D3C0
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblNon 
      BackStyle       =   0  'Transparent
      Caption         =   "Non - Walkable"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image imgTile 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   31
      Left            =   1200
      Picture         =   "frmOptions.frx":1D880
      Top             =   1680
      Width           =   435
   End
   Begin VB.Image imgSprite 
      Height          =   1275
      Index           =   15
      Left            =   2160
      Picture         =   "frmOptions.frx":1DC69
      Top             =   5400
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgSprite 
      Height          =   1650
      Index           =   14
      Left            =   120
      Picture         =   "frmOptions.frx":1ED40
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgSprite 
      Height          =   1035
      Index           =   13
      Left            =   6600
      Picture         =   "frmOptions.frx":20002
      Top             =   3240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgSprite 
      Height          =   2490
      Index           =   12
      Left            =   5400
      Picture         =   "frmOptions.frx":20520
      Top             =   2280
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   30
      Left            =   2280
      Picture         =   "frmOptions.frx":22980
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   29
      Left            =   1800
      Picture         =   "frmOptions.frx":22D8C
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgSprite 
      Height          =   1620
      Index           =   11
      Left            =   3240
      Picture         =   "frmOptions.frx":23175
      Top             =   3840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgSprite 
      Height          =   1620
      Index           =   10
      Left            =   6960
      Picture         =   "frmOptions.frx":24408
      Top             =   5400
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image imgSprite 
      Height          =   1620
      Index           =   9
      Left            =   3960
      Picture         =   "frmOptions.frx":25232
      Top             =   5400
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image imgSprite 
      Height          =   240
      Index           =   8
      Left            =   1320
      Picture         =   "frmOptions.frx":26063
      Top             =   3720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgSprite 
      Height          =   330
      Index           =   7
      Left            =   960
      Picture         =   "frmOptions.frx":26401
      Top             =   3720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgSprite 
      Height          =   435
      Index           =   6
      Left            =   600
      Picture         =   "frmOptions.frx":267C8
      Top             =   3720
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSprite 
      Height          =   540
      Index           =   5
      Left            =   240
      Picture         =   "frmOptions.frx":26BB6
      Top             =   3720
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgSprite 
      Height          =   750
      Index           =   4
      Left            =   2040
      Picture         =   "frmOptions.frx":26FE1
      Top             =   4320
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgSprite 
      Height          =   675
      Index           =   3
      Left            =   1560
      Picture         =   "frmOptions.frx":275EA
      Top             =   4320
      Width           =   360
   End
   Begin VB.Image imgSprite 
      Height          =   675
      Index           =   2
      Left            =   1080
      Picture         =   "frmOptions.frx":27C16
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSprite 
      Height          =   585
      Index           =   1
      Left            =   600
      Picture         =   "frmOptions.frx":282BD
      Top             =   4440
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgSprite 
      Height          =   675
      Index           =   0
      Left            =   120
      Picture         =   "frmOptions.frx":288E8
      Top             =   4320
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   632
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As Online Town Map"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveMaze 
         Caption         =   "Save As &Maze"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveMazeOff 
         Caption         =   "Save As Official Maze"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveQuest 
         Caption         =   "Save As Single &Player Quest"
      End
      Begin VB.Menu mnuLB4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Town"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpenMaze 
         Caption         =   "Open Ma&ze"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpenMazeOff 
         Caption         =   "Open Official Maze"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpenQuest 
         Caption         =   "Open &Quest"
      End
      Begin VB.Menu mnuLB2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFlood 
         Caption         =   "Flood With Current Tile (New Map)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuLB3 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuQuest 
      Caption         =   "&Quest Options"
      Begin VB.Menu mnuBattle 
         Caption         =   "&Random Battle Editor"
      End
      Begin VB.Menu mnuBoss 
         Caption         =   "&Show Boss Editor"
      End
      Begin VB.Menu mnuOverallQuest 
         Caption         =   "&Show Overall Quest Editor"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkLayer_Click()

If chkLayer.Value = 1 Then
    For i = 1 To 25
        If SpriteType(i) <> 999 Then
            frmEditor.imgSprite(i).Visible = True
        End If
    Next 'i
End If
If chkLayer.Value = 0 Then
    For i = 1 To 25
        frmEditor.imgSprite(i).Visible = False
        If SpriteType(i) = 17 Then
        frmEditor.Cls
            'Call BitBlt(frmEditor.hDC, imgSprite(i).Left, imgSprite(i).Top, imgSprite(i).Width, imgSprite(i).Height, picHouse.hDC, 0, 0, vbSrcAnd)
            'Call BitBlt(frmEditor.hDC, imgSprite(i).Left, imgSprite(i).Top, imgSprite(i).Width, imgSprite(i).Height, picHouse.hDC, 0, 0, vbSrcPaint)
        frmEditor.Refresh
        End If
    Next 'i
End If

End Sub

Private Sub Form_Load()
curType = 0
End Sub

Private Sub imgSprite_Click(Index As Integer)
CurTile = Index
curType = 1
End Sub

Private Sub imgTile_Click(Index As Integer)
CurTile = Index
curType = 0
End Sub

Private Sub mnuBattle_Click()
frmBattle.Show
frmBattle.framBoss.Visible = False
frmBattle.framRandom.Visible = True
End Sub

Private Sub mnuBoss_Click()
frmBattle.Show
frmBattle.framBoss.Visible = True
frmBattle.framRandom.Visible = False
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuFlood_Click()
On Error Resume Next
    For i = 1 To 280
        frmEditor.imgTile(i).Picture = frmOptions.imgTile(CurTile).Picture
        If CurTile = 31 Then
            frmEditor.imgTile(i).BorderStyle = 1
        Else
            frmEditor.imgTile(i).BorderStyle = 0
        End If
        
        If frmOptions.chkGoTo.Value = 1 Then
            TileGoto(i) = True
            TileLink(i) = frmOptions.txtLevel.Text
            frmEditor.imgTile(i).BorderStyle = 1
            TileGotoLink(i) = frmOptions.txtGoto.Text
        Else
            TileGoto(i) = False
            If CurTile <> 31 Then
                frmEditor.imgTile(i).BorderStyle = 0
            End If
        End If
    
        If frmOptions.chkRandom.Enabled = True Then
            TileRandom(i) = True
        Else
            TileRandom(i) = False
        End If
  '      If frmTalk.chkTalk.Value = 1 Then
  '          TileTalk(i) = frmTalk.txtTalk.Text
  '      Else
  '          TileTalk(i) = ""
  '      End If
        If frmOptions.chkBoss.Value = 1 Then
            TileBoss(i) = True
        Else
            TileBoss(i) = False
        End If
    If frmOptions.chkJump.Value = 1 Then
        TileJumpable(i) = True
    Else
        TileJumpable(i) = False
    End If
    If frmOptions.chkMove.Value = 1 Then
        TileMovable(i) = True
    Else
        TileMovable(i) = False
    End If
    TileType(i) = CurTile
    Next 'i
End Sub

Private Sub mnuHelp_Click()
frmHelp.Show
End Sub

Private Sub mnuOpen_Click()
On Error Resume Next
Dim strSave As String
strSave = InputBox("Enter name of file to open w/out extension")
strSave = App.Path & "\" & strSave & ".dat"
For i = 1 To 280
    Dim intCur As Integer
    intCur = CInt(GetFromIni("GEN", "T" & i, strSave))
    TileType(i) = intCur
    frmEditor.imgTile(i).Picture = imgTile(TileType(i)).Picture
    If TileType(i) = 31 Then
        frmEditor.imgTile(i).BorderStyle = 1
    Else
        frmEditor.imgTile(i).BorderStyle = 0
    End If
    
    Dim strTileLink As String
    strTileLink = GetFromIni("GEN", "G" & i, strSave)
    
    If strTileLink = "T" Then
        TileGoto(i) = True
        TileLink(i) = GetFromIni("GEN", "GL" & i, strSave)
    Else
        TileGoto(i) = False
    End If
    If TileGoto(i) = True Then
        frmEditor.imgTile(i).BorderStyle = 1
    End If
    
    
    If i <= 25 Then
        SpriteType(i) = CInt(GetFromIni("GEN", "STYPE" & i, strSave))
        frmEditor.imgSprite(i).Picture = imgSprite(SpriteType(i)).Picture
        frmEditor.imgSprite(i).Left = CInt(GetFromIni("GEN", "SLEFT" & i, strSave))
        frmEditor.imgSprite(i).Top = CInt(GetFromIni("GEN", "STOP" & i, strSave))
        If SpriteType(i) <> 999 Then
            frmEditor.imgSprite(i).Visible = True
        Else
            frmEditor.imgSprite(i).Visible = False
        End If
    End If
Next 'i
End Sub

Private Sub mnuOpenMaze_Click()
On Error GoTo err
Dim strSave As String
'strSave = InputBox("Enter full path and file name of file to open w/out extension.  For example: 'C:\Program Files\War of the Adepts\NewMaze'")
'strSave = strSave & ".maz"
comDiag.Filter = "War of the Adepts Maze File (*.maz)|*.maz"
comDiag.ShowOpen
'If comDiag.CancelError = False Then Exit Sub
strSave = comDiag.FileName
txtMazeStart.Text = GetFromIni("GEN", "START", strSave)

For i = 1 To 280
    Dim intCur As Integer
    intCur = CInt(GetFromIni("GEN", "T" & i, strSave))
    TileType(i) = intCur
    frmEditor.imgTile(i).Picture = imgTile(TileType(i)).Picture
    If TileType(i) = 31 Then
        frmEditor.imgTile(i).BorderStyle = 1
    Else
        frmEditor.imgTile(i).BorderStyle = 0
    End If
    
    Dim strTemp As String
    strTemp = GetFromIni("GEN", "F" & i, strSave)
    If strTemp = "1" Then
        bMazeFinish(i) = 1
        frmEditor.imgTile(i).BorderStyle = 1
    Else
        bMazeFinish(i) = 0
        frmEditor.imgTile(i).BorderStyle = 0
    End If
        
    If i <= 25 Then
        SpriteType(i) = CInt(GetFromIni("GEN", "STYPE" & i, strSave))
        frmEditor.imgSprite(i).Picture = imgSprite(SpriteType(i)).Picture
        frmEditor.imgSprite(i).Left = CInt(GetFromIni("GEN", "SLEFT" & i, strSave))
        frmEditor.imgSprite(i).Top = CInt(GetFromIni("GEN", "STOP" & i, strSave))
        If SpriteType(i) <> 999 Then
            frmEditor.imgSprite(i).Visible = True
        Else
            frmEditor.imgSprite(i).Visible = False
        End If
    End If
Next 'i
Exit Sub
err:
If err.Number = 32755 Then
    Exit Sub
Else
    Debug.Print err.Number
    Resume Next
End If

End Sub

Private Sub mnuOpenMazeOff_Click()
On Error GoTo err
'strSave = InputBox("Enter full path and file name of file to open w/out extension.  For example: 'C:\Program Files\War of the Adepts\NewMaze'")
'strSave = strSave & ".maz"
Dim strSave As String
comDiag.Filter = "WOTA Official Maze File (*.omaz)|*.omaz"
comDiag.ShowOpen
'If comDiag.CancelError = False Then Exit Sub
strSave = comDiag.FileName
txtMazeStart.Text = UltraDecode("START", "STARTL", strSave)

For i = 1 To 280
    Dim intCur As Integer
    intCur = CInt(UltraDecode("T" & CStr(i), "TL" & CStr(i), strSave))
    TileType(i) = intCur
    frmEditor.imgTile(i).Picture = imgTile(TileType(i)).Picture
    If TileType(i) = 31 Then
        frmEditor.imgTile(i).BorderStyle = 1
    Else
        frmEditor.imgTile(i).BorderStyle = 0
    End If
    
    Dim strTemp As String
    strTemp = UltraDecode("F" & CStr(i), "FL" & CStr(i), strSave)
    If strTemp = "1" Then
        bMazeFinish(i) = 1
        frmEditor.imgTile(i).BorderStyle = 1
    Else
        bMazeFinish(i) = 0
        frmEditor.imgTile(i).BorderStyle = 0
    End If
        
    If i <= 25 Then
        SpriteType(i) = CInt(UltraDecode("STYPE" & CStr(i), "STYPEL" & CStr(i), strSave))
        frmEditor.imgSprite(i).Picture = imgSprite(SpriteType(i)).Picture
        frmEditor.imgSprite(i).Left = CInt(UltraDecode("SLEFT" & CStr(i), "SLEFTL" & CStr(i), strSave))
        frmEditor.imgSprite(i).Top = CInt(UltraDecode("STOP" & CStr(i), "STOPL" & CStr(i), strSave))
        If SpriteType(i) <> 999 Then
            frmEditor.imgSprite(i).Visible = True
        Else
            frmEditor.imgSprite(i).Visible = False
        End If
    End If
Next 'i
Exit Sub
err:
If err.Number = 32755 Then
    Exit Sub
Else
    Debug.Print err.Number
    Resume Next
End If
End Sub

Private Sub mnuOpenQuest_Click()
On Error Resume Next
Dim strSave As String
strSave = InputBox("Enter name of file to open w/out extension")
If strSave = "" Then
    Exit Sub
End If
strSave = App.Path & "\" & strSave & ".dat"
For i = 1 To 280
    Dim intCur As Integer
    intCur = CInt(UltraDecode("T" & i, "TL" & i, strSave))
    TileType(i) = intCur
    frmEditor.imgTile(i).Picture = imgTile(TileType(i)).Picture
    
    If TileType(i) = 31 Then
        frmEditor.imgTile(i).BorderStyle = 1
    Else
        frmEditor.imgTile(i).BorderStyle = 0
    End If
    
    Dim strTileLink As String
    strTileLink = UltraDecode("GLVL" & i, "GLVLL" & i, strSave)
    
    If strTileLink <> "NULL" Then
        TileGoto(i) = True
        TileLink(i) = strTileLink
        TileGotoLink(i) = UltraDecode("GX" & i, "GXL" & i, strSave)
    Else
        TileGoto(i) = False
    End If
    
    If TileGoto(i) = True Then
        frmEditor.imgTile(i).BorderStyle = 1
    End If
    
    Dim strBoss As String
    strBoss = UltraDecode("BO" & i, "BOL" & i, strSave)
    If strBoss = "T" Then
        TileBoss(i) = True
    Else
        TileBoss(i) = False
    End If
    
'    Dim strJump As String
'    strJump = UltraDecode("J" & i, "JL" & i, strSave)
'    If strJump = "T" Then
'        TileJumpable(i) = True
'    Else
'        TileJumpable(i) = False
'    End If
    
'    strJump = UltraDecode("MV" & i, "MVL" & i, strSave)
'    If strJump = "T" Then
'        TileMovable(i) = True
'    Else
'        TileMovable(i) = False
'    End If
    
    
    If i <= 25 Then
        frmEditor.imgSprite(i).Visible = False
        SpriteType(i) = CInt(GetFromIni("GEN", "STYPE" & i, strSave))
        frmEditor.imgSprite(i).Picture = imgSprite(SpriteType(i)).Picture
        frmEditor.imgSprite(i).Left = CInt(GetFromIni("GEN", "SLEFT" & i, strSave))
        frmEditor.imgSprite(i).Top = CInt(GetFromIni("GEN", "STOP" & i, strSave))
        If SpriteType(i) <> 999 Then
            frmEditor.imgSprite(i).Visible = True
        Else
            frmEditor.imgSprite(i).Visible = False
        End If
    End If
    
    If i <= 4 Then
        frmBattle.txtAI(i).Text = UltraDecode("BAI" & i, "BAIL" & i, strSave)
        frmBattle.txtAP(i).Text = UltraDecode("BAP" & i, "BAPL" & i, strSave)
        frmBattle.txtHP(i).Text = UltraDecode("BHP" & i, "BHPL" & i, strSave)
        frmBattle.txtDefense(i).Text = UltraDecode("BDEF" & i, "BDEFL" & i, strSave)
        frmBattle.txtCoins(i).Text = UltraDecode("BCOINS" & i, "BCOINSL" & i, strSave)
        frmBattle.txtName(i).Text = UltraDecode("BNAME" & i, "BNAMEL" & i, strSave)
        frmBattle.txtPicture(i).Text = UltraDecode("BPIC" & i, "BPICL" & i, strSave)
    End If
    
    If i = 4 Then
        frmBattle.txtNextMap.Text = UltraDecode("BNEXT" & i, "BNEXTL" & i, strSave)
        frmBattle.txtNextTile.Text = UltraDecode("BNT" & i, "BNTL" & i, strSave)
        frmBattle.txtTalk.Text = UltraDecode("BTALK" & i, "BTALKL" & i, strSave)
        frmBattle.txtTalk2.Text = UltraDecode("BTALK2" & i, "BTALK2L" & i, strSave)
    End If
    
    Dim strTemp As String
    strTemp = UltraDecode("RN" & i, "RNL" & i, strSave)
    If strTemp = "T" Then
        TileRandom(i) = True
    Else
        TileRandom(i) = False
    End If
    
Next 'i
End Sub

Private Sub mnuOverallQuest_Click()
frmQuest.Show
End Sub

Private Sub mnuSave_Click()
Dim strSave As String
strSave = InputBox("Enter name of file and where to save them to  w/out extension")
strSave = App.Path & "\" & strSave & ".dat"
For i = 1 To 280
    Call WriteIni("GEN", "T" & i, CStr(TileType(i)), strSave)
    If TileGoto(i) = True Then
        Call WriteIni("GEN", "G" & i, "T", strSave)
        Call WriteIni("GEN", "GL" & i, TileLink(i), strSave)
    Else
        Call WriteIni("GEN", "G" & i, "F", strSave)
    End If
    If i <= 25 Then
        Call WriteIni("GEN", "STYPE" & i, CStr(SpriteType(i)), strSave)
        Call WriteIni("GEN", "SLEFT" & i, CStr(frmEditor.imgSprite(i).Left), strSave)
        Call WriteIni("GEN", "STOP" & i, CStr(frmEditor.imgSprite(i).Top), strSave)
    End If
Next 'i
End Sub

Private Sub mnuSaveMaze_Click()
On Error Resume Next
Dim strSave As String
'strSave = InputBox("Enter the full path of where to save the file without the extension.  For example, 'C:\Program Files\War of the Adepts\NewMaze'")
'strSave = strSave & ".maz"
comDiag.ShowSave
'If comDiag.CancelError = False Then
'    Exit Sub
'End If

strSave = comDiag.FileName

If Right$(strSave, 4) = "omaz" Then
    MsgBox "Error: Cannot save as an official maze file."
    Exit Sub
End If

Call WriteIni("GEN", "START", txtMazeStart.Text, strSave)

For i = 1 To 280
    Call WriteIni("GEN", "T" & i, CStr(TileType(i)), strSave)
    If bMazeFinish(i) = 1 Then
        Call WriteIni("GEN", "F" & i, "1", strSave)
    Else
        Call WriteIni("GEN", "F" & i, "0", strSave)
    End If
    If i <= 25 Then
        Call WriteIni("GEN", "STYPE" & i, CStr(SpriteType(i)), strSave)
        Call WriteIni("GEN", "SLEFT" & i, CStr(frmEditor.imgSprite(i).Left), strSave)
        Call WriteIni("GEN", "STOP" & i, CStr(frmEditor.imgSprite(i).Top), strSave)
    End If
Next 'i
End Sub

Private Sub mnuSaveMazeOff_Click()
On Error Resume Next
comDiag.Filter = "Official WOTA Maze File (*.omaz) |*.omaz"
comDiag.DefaultExt = ".omaz"
comDiag.ShowSave
Dim nSave As String
nSave = comDiag.FileName

Call Encode(txtMazeStart.Text, "START", "STARTL", nSave)

For i = 1 To 280
    Call Encode(CStr(TileType(i)), "T" & CStr(i), "TL" & CStr(i), nSave)
    If bMazeFinish(i) = 1 Then
        Call Encode("1", "F" & CStr(i), "FL" & CStr(i), nSave)
    Else
        Call Encode("0", "F" & CStr(i), "FL" & CStr(i), nSave)
    End If
    If i <= 25 Then
        Call Encode(CStr(SpriteType(i)), "STYPE" & CStr(i), "STYPEL" & CStr(i), nSave)
        Call Encode(CStr(frmEditor.imgSprite(i).Left), "SLEFT" & CStr(i), "SLEFTL" & CStr(i), nSave)
        Call Encode(CStr(frmEditor.imgSprite(i).Top), "STOP" & CStr(i), "STOPL" & CStr(i), nSave)
    End If
Next 'i

End Sub

Private Sub mnuSaveQuest_Click()
On Error Resume Next

Dim strSave As String
Dim xSave As String
strSave = InputBox("Enter name of file to save w/out extension")
xSave = strSave
strSave = App.Path & "\" & strSave & ".dat"
For i = 1 To 280

    Call Encode(CStr(TileType(i)), "T" & i, "TL" & i, strSave)
    If TileGoto(i) = True Then
    Call Encode(CStr(TileLink(i)), "GLVL" & i, "GLVLL" & i, strSave)
    Call Encode(CStr(TileGotoLink(i)), "GX" & i, "GXL" & i, strSave)
    
    Else
        Call Encode("NULL", "GLVL" & i, "GLVLL" & i, strSave)
    End If
    
    If TileRandom(i) = True Then
        Call Encode("T", "RN" & i, "RNL" & i, strSave)
    Else
        Call Encode("F", "RN" & i, "RNL" & i, strSave)
    End If
    
    If TileBoss(i) = True Then
        Call Encode("T", "BO" & i, "BOL" & i, strSave)
    Else
        Call Encode("F", "BO" & i, "BOL" & i, strSave)
    End If
    
'    If TileJumpable(i) = True Then
'        Call Encode("T", "J" & i, "JL" & i, strSave)
'    Else
'        Call Encode("F", "J" & i, "JL" & i, strSave)
'    End If
'
'    If TileMovable(i) = True Then
'        Call Encode("T", "MV" & i, "MVL" & i, strSave)
'    Else
'        Call Encode("F", "MV" & i, "MVL" & i, strSave)
'    End If

    If i <= 4 Then
        Call Encode(frmBattle.txtHP(i).Text, "BHP" & i, "BHPL" & i, strSave)
        Call Encode(frmBattle.txtAP(i).Text, "BAP" & i, "BAPL" & i, strSave)
        Call Encode(frmBattle.txtDefense(i).Text, "BDEF" & i, "BDEFL" & i, strSave)
        Call Encode(frmBattle.txtAI(i).Text, "BAI" & i, "BAIL" & i, strSave)
        Call Encode(frmBattle.txtPicture(i).Text, "BPIC" & i, "BPICL" & i, strSave)
        Call Encode(frmBattle.txtCoins(i).Text, "BCOINS" & i, "BCOINSL" & i, strSave)
        Call Encode(frmBattle.txtName(i).Text, "BNAME" & i, "BNAMEL" & i, strSave)
    End If
    If i = 4 Then
        Call Encode(frmBattle.txtNextMap.Text, "BNEXT" & i, "BNEXTL" & i, strSave)
        Call Encode(frmBattle.txtNextTile.Text, "BNT" & i, "BNTL" & i, strSave)
        Call Encode(frmBattle.txtTalk.Text, "BTALK" & i, "BTALKL" & i, strSave)
        Call Encode(frmBattle.txtTalk.Text, "BTALK2" & i, "BTALK2L" & i, strSave)
    End If
    
    
    
    If i <= 25 Then
        Call WriteIni("GEN", "STYPE" & i, CStr(SpriteType(i)), strSave)
        Call WriteIni("GEN", "SLEFT" & i, CStr(frmEditor.imgSprite(i).Left), strSave)
        Call WriteIni("GEN", "STOP" & i, CStr(frmEditor.imgSprite(i).Top), strSave)
    End If
Next 'i
If Admin = False Then
    Call Encode("FALSE", "OFFICIAL", "OFFICIALL", strSave)
Else
    Call Encode(xSave, "OFFICIAL", "OFFICIALL", strSave)
End If

End Sub


Private Sub mnuUpload_Click()

End Sub
