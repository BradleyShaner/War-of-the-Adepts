VERSION 5.00
Begin VB.Form frmArena 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "War of the Adepts Pre-Battle Maze"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   ControlBox      =   0   'False
   Icon            =   "frmArena.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   20
      Left            =   2640
      Picture         =   "frmArena.frx":08CA
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   76
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   19
      Left            =   2160
      Picture         =   "frmArena.frx":0C70
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   75
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   18
      Left            =   1680
      Picture         =   "frmArena.frx":1018
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   74
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   20
      Left            =   2520
      Picture         =   "frmArena.frx":13B4
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   73
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   19
      Left            =   2040
      Picture         =   "frmArena.frx":175A
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   72
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   18
      Left            =   1560
      Picture         =   "frmArena.frx":1B02
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   71
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   20
      Left            =   2760
      Picture         =   "frmArena.frx":1E9E
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   70
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   19
      Left            =   1440
      Picture         =   "frmArena.frx":24E4
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   69
      Top             =   3000
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   20
      Left            =   2640
      Picture         =   "frmArena.frx":2A8C
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   68
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   19
      Left            =   1440
      Picture         =   "frmArena.frx":2D77
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   67
      Top             =   2880
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   17
      Left            =   0
      Picture         =   "frmArena.frx":2FBA
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   66
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   17
      Left            =   3000
      Picture         =   "frmArena.frx":3363
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   65
      Top             =   3360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   12
      Left            =   2640
      Picture         =   "frmArena.frx":354F
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   64
      Top             =   3360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   17
      Left            =   3120
      Picture         =   "frmArena.frx":373B
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   63
      Top             =   4320
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   12
      Left            =   2760
      Picture         =   "frmArena.frx":3E19
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   62
      Top             =   4320
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1710
      Index           =   18
      Left            =   480
      Picture         =   "frmArena.frx":44F7
      ScaleHeight     =   114
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   60
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1710
      Index           =   18
      Left            =   480
      Picture         =   "frmArena.frx":4E47
      ScaleHeight     =   114
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   59
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Index           =   16
      Left            =   1920
      Picture         =   "frmArena.frx":5AD1
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   58
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Index           =   16
      Left            =   1920
      Picture         =   "frmArena.frx":60E9
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   57
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   56
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1275
      Index           =   15
      Left            =   5760
      Picture         =   "frmArena.frx":61D5
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   55
      Top             =   3000
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1650
      Index           =   14
      Left            =   3600
      Picture         =   "frmArena.frx":6381
      ScaleHeight     =   110
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   54
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1275
      Index           =   15
      Left            =   5640
      Picture         =   "frmArena.frx":714B
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   53
      Top             =   2880
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1650
      Index           =   14
      Left            =   3600
      Picture         =   "frmArena.frx":821A
      ScaleHeight     =   110
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   52
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   13
      Left            =   1920
      Picture         =   "frmArena.frx":94D4
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   51
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1620
      Index           =   11
      Left            =   3600
      Picture         =   "frmArena.frx":9C5E
      ScaleHeight     =   108
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   50
      Top             =   3600
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1620
      Index           =   10
      Left            =   5880
      Picture         =   "frmArena.frx":B0DB
      ScaleHeight     =   108
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   49
      Top             =   3600
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1620
      Index           =   9
      Left            =   3840
      Picture         =   "frmArena.frx":C14F
      ScaleHeight     =   108
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   48
      Top             =   3600
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   13
      Left            =   1920
      Picture         =   "frmArena.frx":D1BE
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   47
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1620
      Index           =   11
      Left            =   3720
      Picture         =   "frmArena.frx":D601
      ScaleHeight     =   108
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   46
      Top             =   3480
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1620
      Index           =   10
      Left            =   5760
      Picture         =   "frmArena.frx":D826
      ScaleHeight     =   108
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   45
      Top             =   3600
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1620
      Index           =   9
      Left            =   5760
      Picture         =   "frmArena.frx":E40E
      ScaleHeight     =   108
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   44
      Top             =   3600
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Timer timeCount 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   4800
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   16
      Left            =   960
      Picture         =   "frmArena.frx":EFF6
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   43
      Top             =   840
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   15
      Left            =   240
      Picture         =   "frmArena.frx":F072
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   42
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   14
      Left            =   6600
      Picture         =   "frmArena.frx":F0F7
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   41
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   13
      Left            =   6240
      Picture         =   "frmArena.frx":F181
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   40
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   12
      Left            =   5880
      Picture         =   "frmArena.frx":F200
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   11
      Left            =   5400
      Picture         =   "frmArena.frx":F27C
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   38
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   10
      Left            =   4800
      Picture         =   "frmArena.frx":F2FC
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   9
      Left            =   4440
      Picture         =   "frmArena.frx":F388
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   36
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   8
      Left            =   3960
      Picture         =   "frmArena.frx":F406
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   7
      Left            =   3480
      Picture         =   "frmArena.frx":F488
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   34
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   6
      Left            =   3000
      Picture         =   "frmArena.frx":F50B
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   5
      Left            =   2640
      Picture         =   "frmArena.frx":F591
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   32
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   4
      Left            =   2160
      Picture         =   "frmArena.frx":F610
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   31
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   3
      Left            =   1680
      Picture         =   "frmArena.frx":F690
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   2
      Left            =   1200
      Picture         =   "frmArena.frx":F712
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   29
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   1
      Left            =   840
      Picture         =   "frmArena.frx":F792
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellowM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   0
      Left            =   360
      Picture         =   "frmArena.frx":F816
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   16
      Left            =   960
      Picture         =   "frmArena.frx":F896
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   15
      Left            =   480
      Picture         =   "frmArena.frx":FC34
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   14
      Left            =   6840
      Picture         =   "frmArena.frx":FFDD
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   13
      Left            =   6360
      Picture         =   "frmArena.frx":1038D
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   12
      Left            =   5880
      Picture         =   "frmArena.frx":1072E
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   11
      Left            =   5400
      Picture         =   "frmArena.frx":10ACC
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   10
      Left            =   4800
      Picture         =   "frmArena.frx":10E6E
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   9
      Left            =   4320
      Picture         =   "frmArena.frx":11220
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   8
      Left            =   3840
      Picture         =   "frmArena.frx":115C0
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   7
      Left            =   3360
      Picture         =   "frmArena.frx":11965
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   6
      Left            =   3000
      Picture         =   "frmArena.frx":11D0C
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   5
      Left            =   2640
      Picture         =   "frmArena.frx":120B6
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   4
      Left            =   2280
      Picture         =   "frmArena.frx":12457
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   3
      Left            =   1800
      Picture         =   "frmArena.frx":127F9
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   2
      Left            =   1320
      Picture         =   "frmArena.frx":12B9E
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   1
      Left            =   840
      Picture         =   "frmArena.frx":12F40
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   0
      Left            =   480
      Picture         =   "frmArena.frx":132E8
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   4
      Left            =   1920
      Picture         =   "frmArena.frx":1368A
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   3
      Left            =   1440
      Picture         =   "frmArena.frx":13833
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   2
      Left            =   960
      Picture         =   "frmArena.frx":139EC
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   1
      Left            =   240
      Picture         =   "frmArena.frx":13B9E
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   720
      Picture         =   "frmArena.frx":13D5B
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   4
      Left            =   1920
      Picture         =   "frmArena.frx":13F47
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   3
      Left            =   1440
      Picture         =   "frmArena.frx":14549
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   2
      Left            =   960
      Picture         =   "frmArena.frx":14B75
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   1
      Left            =   480
      Picture         =   "frmArena.frx":15212
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   0
      Picture         =   "frmArena.frx":1583D
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   67
      Left            =   5640
      Picture         =   "frmArena.frx":15F1B
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   68
      Left            =   5640
      Picture         =   "frmArena.frx":16331
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   66
      Left            =   5880
      Picture         =   "frmArena.frx":16746
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   38
      Left            =   7200
      Picture         =   "frmArena.frx":16B6B
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblYouWin 
      BackStyle       =   0  'Transparent
      Caption         =   "You Win!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   61
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   64
      Left            =   6240
      Picture         =   "frmArena.frx":16FEB
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   65
      Left            =   6600
      Picture         =   "frmArena.frx":1742B
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   59
      Left            =   6960
      Picture         =   "frmArena.frx":1786A
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   36
      Left            =   5160
      Picture         =   "frmArena.frx":17CEA
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   37
      Left            =   5160
      Picture         =   "frmArena.frx":18175
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   57
      Left            =   5280
      Picture         =   "frmArena.frx":185F0
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   58
      Left            =   5280
      Picture         =   "frmArena.frx":18A35
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   54
      Left            =   6720
      Picture         =   "frmArena.frx":18E8D
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   55
      Left            =   6600
      Picture         =   "frmArena.frx":192E6
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   56
      Left            =   7080
      Picture         =   "frmArena.frx":1970E
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   60
      Left            =   6600
      Picture         =   "frmArena.frx":19B61
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   61
      Left            =   6960
      Picture         =   "frmArena.frx":19FA5
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   62
      Left            =   6600
      Picture         =   "frmArena.frx":1A3E0
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   63
      Left            =   6600
      Picture         =   "frmArena.frx":1A805
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   42
      Left            =   6720
      Picture         =   "frmArena.frx":1AC2A
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   43
      Left            =   6720
      Picture         =   "frmArena.frx":1B0D7
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   44
      Left            =   6720
      Picture         =   "frmArena.frx":1B595
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   45
      Left            =   6720
      Picture         =   "frmArena.frx":1BA21
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   46
      Left            =   7080
      Picture         =   "frmArena.frx":1BEE2
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   47
      Left            =   7080
      Picture         =   "frmArena.frx":1C38E
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   48
      Left            =   4560
      Picture         =   "frmArena.frx":1C821
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   49
      Left            =   4560
      Picture         =   "frmArena.frx":1CCB9
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   50
      Left            =   6240
      Picture         =   "frmArena.frx":1D154
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   51
      Left            =   6240
      Picture         =   "frmArena.frx":1D5E5
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   52
      Left            =   6600
      Picture         =   "frmArena.frx":1DA97
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   53
      Left            =   6600
      Picture         =   "frmArena.frx":1DF2C
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   32
      Left            =   5880
      Picture         =   "frmArena.frx":1E3DB
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   33
      Left            =   6240
      Picture         =   "frmArena.frx":1E89B
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   34
      Left            =   5880
      Picture         =   "frmArena.frx":1ED3E
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   35
      Left            =   6240
      Picture         =   "frmArena.frx":1F226
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   39
      Left            =   7200
      Picture         =   "frmArena.frx":1F6EB
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   40
      Left            =   7920
      Picture         =   "frmArena.frx":1FB65
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   41
      Left            =   7920
      Picture         =   "frmArena.frx":1FEB6
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   31
      Left            =   6960
      Picture         =   "frmArena.frx":202A5
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   0
      Left            =   3480
      Picture         =   "frmArena.frx":2068E
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   1
      Left            =   3360
      Picture         =   "frmArena.frx":20BCF
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   2
      Left            =   2880
      Picture         =   "frmArena.frx":210AF
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   3
      Left            =   3960
      Picture         =   "frmArena.frx":2161D
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   4
      Left            =   2280
      Picture         =   "frmArena.frx":21B77
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   5
      Left            =   4680
      Picture         =   "frmArena.frx":22137
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   6
      Left            =   4200
      Picture         =   "frmArena.frx":226EA
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   7
      Left            =   4920
      Picture         =   "frmArena.frx":22C84
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   8
      Left            =   5640
      Picture         =   "frmArena.frx":23218
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   9
      Left            =   4080
      Picture         =   "frmArena.frx":23792
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   10
      Left            =   4800
      Picture         =   "frmArena.frx":23D31
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   11
      Left            =   5520
      Picture         =   "frmArena.frx":242CE
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   12
      Left            =   3000
      Picture         =   "frmArena.frx":2484F
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   13
      Left            =   4200
      Picture         =   "frmArena.frx":24DC7
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   14
      Left            =   5640
      Picture         =   "frmArena.frx":25320
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   15
      Left            =   3840
      Picture         =   "frmArena.frx":25875
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   16
      Left            =   1800
      Picture         =   "frmArena.frx":25DF2
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   17
      Left            =   5520
      Picture         =   "frmArena.frx":2635D
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   18
      Left            =   6240
      Picture         =   "frmArena.frx":2690E
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   19
      Left            =   6120
      Picture         =   "frmArena.frx":26EC7
      Top             =   -120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   20
      Left            =   6600
      Picture         =   "frmArena.frx":27453
      Top             =   -120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   21
      Left            =   2400
      Picture         =   "frmArena.frx":279E3
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   22
      Left            =   3120
      Picture         =   "frmArena.frx":27E28
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   23
      Left            =   3840
      Picture         =   "frmArena.frx":28245
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   24
      Left            =   4560
      Picture         =   "frmArena.frx":2864D
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   25
      Left            =   5280
      Picture         =   "frmArena.frx":28A7E
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   26
      Left            =   6000
      Picture         =   "frmArena.frx":28EB5
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   27
      Left            =   6240
      Picture         =   "frmArena.frx":292FE
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   28
      Left            =   6840
      Picture         =   "frmArena.frx":29880
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   29
      Left            =   7800
      Picture         =   "frmArena.frx":29DFB
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   30
      Left            =   8280
      Picture         =   "frmArena.frx":2A1E4
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgPlat 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmArena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim LastTick, CurrentTick As Long
Dim TickDif As Long
Dim TimeLeft As Integer

Dim bWalkThroughWalls As Boolean
Dim bSpeed As Boolean

'Variables for moving logs:
Dim LogMove(1 To 25) As Byte
Dim bLogHit(1 To 25) As Boolean

Dim curRun As Integer

Dim IsaacX As Integer
Dim IsaacY As Integer

Dim DirLeft As Boolean
Dim DirUp As Boolean
Dim DirRight As Boolean
Dim DirDown As Boolean

Dim DrawTick As Integer

Dim bRunning As Boolean

Dim Isaac As GameTile
Dim Sprite(1 To 25) As GameTile
Dim Plat(1 To 280) As GameTile


Private Sub Form_Activate()
On Error Resume Next
'If BattleLoaded(1) = True Then
'    timecount.Enabled = False
'    Unload Me
'    frmBattle.Show
'    Exit Sub
'End If
bWalkThroughWalls = False

DrawTick = 0
TickDif = 45 'Set interval for running
intOpMazeX = 0
intOpMazeY = 0

bRunning = True 'Game is running

'You're not moving in any direction
DirUp = False
DirDown = False
DirLeft = False
DirRight = False

'Get rid of the town window
Unload frmMultiplayer

If bMazeFirstLoad = True Then
    Call Form_Load
End If

Call GameLoop


End Sub

Private Sub Form_GotFocus()
If bMazeFirstLoad = True Then
    Call Form_Load
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyEscape And Shift = 1 Then
    Unload Me
End If

If KeyCode = vbKeyLeft And DirLeft = False Then
    If Isaac.Left >= 0 Then
        DirLeft = True
        Isaac.Num = 5
    Else
        DirLeft = False
        Isaac.Num = 15
    End If
End If
If KeyCode = vbKeyUp And DirUp = False Then
    If Isaac.Top >= -15 Then
        DirUp = True
        Isaac.Num = 18
    Else
        DirUp = False
        Isaac.Num = 15
    End If
End If
If KeyCode = vbKeyRight And DirRight = False Then
    If Isaac.Left <= Me.ScaleWidth - Isaac.Width Then
        DirRight = True
        Isaac.Num = 0
    Else
        DirRight = False
        Isaac.Num = 15
    End If
End If
If KeyCode = vbKeyDown And DirDown = False Then
    If Isaac.Top <= Me.ScaleHeight - Isaac.Height Then
        DirDown = True
        Isaac.Num = 18
    Else
        DirDown = False
        Isaac.Num = 15
    End If
End If
If KeyCode = vbKeyQ Then
    MsgBox "Do you have a question?", vbInformation, "Easter Egg #19"
    Call Encode("19", "EGG19", "EGGL19", App.Path & "\settings.ini")
End If
If Shift = 1 And strMyUserName = "admin" Then
    bSpeed = True
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyLeft Then
DirLeft = False
Isaac.Num = 15
End If
If KeyCode = vbKeyUp Then
DirUp = False
Isaac.Num = 15
End If
If KeyCode = vbKeyRight Then
DirRight = False
Isaac.Num = 15
End If
If KeyCode = vbKeyDown Then
DirDown = False
Isaac.Num = 15
End If
If KeyCode = vbKeyNumpad9 And Shift = 1 Then
    bSpeed = False
End If
End Sub

Private Sub Form_Load()

On Error Resume Next

'If bMazeFirstLoad = False And hoston = True Then Exit Sub

'Form_load has initialized
bMazeFirstLoad = False


'Changed for ladder tournament:
If (frmHost.Host.State <> sckClosed Or frmJoin.Client.State <> sckClosed) Then 'Load the maze


    Dim col As Integer 'Columns
    Dim row As Integer 'Rows
    col = 0
    row = 0
    
    If hoston = True Then
    
        Dim RandInt As Integer
        Dim strMap As String
        Dim strRand As String
        strRand = Format(Now, "ss")
        
        Dim intRand As Long
        'intRand = Int(Rnd * frmHost.filMazes.ListCount)
        'strMap = frmHost.filMazes.List(intRand)
        
        Randomize
        intRand = Int(Rnd * 17) + 1
        strMap = "OfficialMaze" & CStr(intRand) & ".omaz"
 
 
'       If RandInt <= 8 Then
'           RandInt = 1
 '       ElseIf RandInt <= 16 Then
 '           RandInt = 2
 '       ElseIf RandInt <= 24 Then
 '           RandInt = 3
 '       ElseIf RandInt <= 32 Then
 '           RandInt = 4
 '       ElseIf RandInt <= 40 Then
 '           RandInt = 1
 '       ElseIf RandInt <= 48 Then
 '           RandInt = 2
 '       ElseIf RandInt <= 56 Then
 '           RandInt = 3
 '       Else
 '           RandInt = 4
 '       End If
 '
        frmHost.Host.SendData "MAPARENA" & strMap & vbCrLf
        DoEvents
    
        strMap = App.Path & "\files\" & strMap
    
    End If
    
    
    For i = 0 To 280 'Clear the map
        If i >= 1 Then
            Plat(i).Left = 25 * col
            Plat(i).Top = 25 * row
            Plat(i).Num = 0
            
            col = col + 1
            If col = 20 Then
                row = row + 1
                col = 0
            End If
        End If
        
        If i <= imgTile.UBound Then
            Load picTile(i)
            picTile(i).Picture = imgTile(i).Picture
            picTile(i).Width = 25
            picTile(i).Height = 25
            picTile(i).Left = 25 * (i - 1)
            picTile(i).Visible = False
        End If
        
        If i <= 25 And i > 0 Then
            Sprite(i).Num = 999
            Sprite(i).Visible = False
        End If
    Next 'i
    
    'Load the map
        
    If hoston = False Then
        strMap = App.Path & "\files\" & strMaptoLoad 'Load the map sent by the host
    End If
    
    'Set Isaac's position
    Isaac.Visible = True
    Isaac.Num = 15
    Dim intStartTile As Long
    intStartTile = CLng(MazeDecode("START", "STARTL", strMap))
    Isaac.Left = Plat(intStartTile).Left + 5
    Isaac.Top = Plat(intStartTile).Top - 15
    IsaacX = 2
    IsaacY = 0
    Isaac.Width = picYellow(0).ScaleWidth
    Isaac.Height = picYellow(0).ScaleHeight
    
'    If strMap = "" Then strMap = "LadMaze1"
    
    Call LoadMap(strMap)
    Call DrawMap
    
    'Play the music
    Call PlayMidi("lighthouse", True)
    
    If hoston = False Then
        frmJoin.Client.SendData "LOADARENA" & vbCrLf 'Tell the host you're ready
        bRunning = True
    Else
        bRunning = True
    End If

    TimeLeft = 45 'Start the clock
    timecount.Enabled = True
    
    

    opFinished = False 'Your opponent has not won the maze


Else 'battleloaded = true
    'MsgBox "Error in connecting to your opponent."
    'Unload Me
End If

'Exit Sub

'err:
'If hoston = False And bMazeError = False Then
'    Dim intMazeRand As Long
'    bMazeError = True
'    Randomize
'    intMazeRand = Int(Rnd * 4) + 1
'    strMaptoLoad = "OfficialMaze" & CStr(intMazeRand) & ".omaz"
'    Call Form_Load
'    Debug.Print err.Description
'    Exit Sub
'Else
'    Resume Next
'End If

End Sub

Sub LoadMap(ByVal strSave As String)
On Error Resume Next

'strSave = App.Path & "\" & strSave & ".dat"
For i = 1 To 280
    Dim intCur As Integer
    intCur = CInt(MazeDecode("T" & CStr(i), "TL" & CStr(i), strSave))
    Plat(i).Num = intCur
    'See if it is an exit tile or not:
    intCur = CInt(MazeDecode("F" & CStr(i), "FL" & CStr(i), strSave))
    If intCur = 1 Then
        Plat(i).Screen = 1
    Else
        Plat(i).Screen = 0
    End If
    
    If i <= 25 Then
        Sprite(i).Num = CInt(MazeDecode("STYPE" & CStr(i), "STYPEL" & CStr(i), strSave))
        Sprite(i).Left = CInt(MazeDecode("SLEFT" & CStr(i), "SLEFTL" & CStr(i), strSave))
        Sprite(i).Top = CInt(MazeDecode("STOP" & CStr(i), "STOPL" & CStr(i), strSave))
        Sprite(i).Height = picSprite(Sprite(i).Num).ScaleHeight
        Sprite(i).Width = picSprite(Sprite(i).Num).ScaleWidth
        If Sprite(i).Num <> 999 Then
            Sprite(i).Visible = True
        Else
            Sprite(i).Visible = False
        End If
    End If
        
Next 'i
End Sub
Sub DrawMap()
On Error Resume Next
For i = 1 To 280
    Call BitBlt(Me.hdc, Plat(i).Left, Plat(i).Top, 25, 25, picTile(Plat(i).Num).hdc, 0, 0, vbSrcCopy)
Next 'i
End Sub


Sub DrawSprites()
On Error Resume Next
    For i = 1 To 25
        If Sprite(i).Visible = True Then
        Call BitBlt(Me.hdc, Sprite(i).Left, Sprite(i).Top, Sprite(i).Width, Sprite(i).Height, picSpriteM(Sprite(i).Num).hdc, 0, 0, vbSrcAnd)
        Call BitBlt(Me.hdc, Sprite(i).Left, Sprite(i).Top, Sprite(i).Width, Sprite(i).Height, picSprite(Sprite(i).Num).hdc, 0, 0, vbSrcPaint)
        End If
    Next 'i
End Sub
Sub GameLoop()
On Error Resume Next
'The main loop of the maze, clears and redraws screen each time
Do While bRunning
    If BattleLoaded(1) = True Then
        bRunning = False
        Unload Me
        Exit Sub
    End If
    If bSpeed = True Then
        TickDif = 15
    Else
        TickDif = 45
    End If
    If hoston = False Or (hoston = True And MazeWait = False) Then 'If I'm not the host or if I'm the host and the client has loaded
    CurrentTick = GetTickCount()
        If CurrentTick - LastTick > TickDif Then
            LastTick = CurrentTick
            Me.Cls
            Call DrawMap
            Call DrawSprites
            Call DrawIsaac
            Call DrawText
            Call CheckWin
            Me.Refresh
            DoEvents
        Else
            DoEvents
        End If
    Else
        DoEvents
    End If
Loop
End Sub
Sub DrawIsaac()

On Error Resume Next

Dim Collision As Boolean

Isaac.Height = picYellow(Isaac.Num).ScaleHeight

If DirLeft = True And Isaac.Left >= 0 Then
    Isaac.Left = Isaac.Left - 5
End If
If DirRight = True And Isaac.Left <= Me.ScaleWidth - 25 Then
    Isaac.Left = Isaac.Left + 5
End If
If DirUp = True And Isaac.Top >= -15 Then
    Isaac.Top = Isaac.Top - 5
End If
If DirDown = True And Isaac.Top <= Me.ScaleHeight - Isaac.Height Then
    Isaac.Top = Isaac.Top + 5
End If

Collision = DetectCollision

'Code for log moving:
For i = 1 To 25
    If Sprite(i).Visible = True Then
        With Sprite(i)
            'If Isaac.Left + 5 <= Sprite(i).Left + Sprite(i).Width And Isaac.Left + 10 >= Sprite(i).Left And Isaac.Top + Abs(25 - Isaac.Height) + 4 <= Sprite(i).Top + Sprite(i).Height And (Isaac.Top + Abs(25 - Isaac.Height)) + 21 >= Sprite(i).Top Then
            If Isaac.Left + 4 <= .Left + .Width And Isaac.Left + Isaac.Width >= .Left And Isaac.Top + Isaac.Height - 27 <= .Top + .Height And Isaac.Top + Isaac.Height - 7 >= .Top Then
                If Sprite(i).Num = 19 Or Sprite(i).Num = 20 Or Sprite(i).Num = 21 Then
                    Collide = True
                End If
                If Sprite(i).Num = 19 Then
                    If DirUp = True Then
                        LogMove(i) = 1
                    End If
                    If DirDown = True Then
                        LogMove(i) = 3
                    End If
                End If
                If Sprite(i).Num = 20 Then
                    If DirRight = True Then
                        LogMove(i) = 2
                    End If
                    If DirLeft = True Then
                        LogMove(i) = 4
                    End If
                End If
            End If
        End With
        If LogMove(i) = 1 Then
            Sprite(i).Top = Sprite(i).Top - 5
        End If
        If LogMove(i) = 2 Then
            Sprite(i).Left = Sprite(i).Left + 5
        End If
        If LogMove(i) = 3 Then
            Sprite(i).Top = Sprite(i).Top + 5
        End If
        If LogMove(i) = 4 Then
            Sprite(i).Left = Sprite(i).Left - 5
        End If
        If LogMove(i) = 5 Then
            Sprite(i).Top = Sprite(i).Top - 5
            If Sprite(i).Top Mod 25 = 0 Then
                bLogHit(i) = True
            End If
        End If
        If LogMove(i) = 7 Then
            Sprite(i).Top = Sprite(i).Top + 5
            If Sprite(i).Top Mod 25 = 0 Then
                bLogHit(i) = True
            End If
        End If
        If LogMove(i) = 6 Then
            Sprite(i).Left = Sprite(i).Left + 5
            If Sprite(i).Left Mod 25 = 0 Then
                bLogHit(i) = True
            End If
        End If
        If LogMove(i) = 8 Then
            Sprite(i).Left = Sprite(i).Left - 5
            If Sprite(i).Left Mod 25 = 0 Then
                bLogHit(i) = True
            End If
        End If
        bLogHit(i) = False
        For q = 1 To 280 'Detect if the log has hit any sprites
            If LogMove(i) > 0 And LogMove(i) < 5 Then
                If Sprite(i).Left < Plat(q).Left + 25 And Sprite(i).Left + Sprite(i).Width > Plat(q).Left And Sprite(i).Top < Plat(q).Top + 25 And Sprite(i).Top + Sprite(i).Height > Plat(q).Top Then
                    With Plat(q)
                        If .Num <> 54 And .Num <> 55 And .Num <> 56 And .Num <> 1 And .Num <> 4 And .Num <> 5 And .Num <> 19 And .Num <> 20 And .Num <> 9 And .Num <> 10 And .Num <> 11 And .Num <> 27 And .Num <> 28 And .Num <> 21 And .Num <> 22 And .Num <> 23 And .Num <> 29 And .Num <> 30 And .Num <> 41 Then
                            bLogHit(i) = True
                        End If 'If .num <> 54...
                    End With ' Plat(i)
                End If 'If sprite(i).left...
            End If ' If LogMove(i) > 0
            
            If q <= 25 And q <> i Then
                If Sprite(q).Visible = True And Sprite(i).Left + 4 <= Sprite(q).Left + Sprite(q).Width And Sprite(i).Left + Sprite(i).Width - 4 >= Sprite(q).Left And Sprite(i).Top + 4 <= Sprite(q).Top + Sprite(q).Height And Sprite(i).Top + Sprite(i).Height - 4 >= Sprite(q).Top Then
                    bLogHit(i) = True
                End If
            End If
            
            If bLogHit(i) = True Then 'If the log has hit
                If LogMove(i) = 1 Then
                    Sprite(i).Top = Sprite(i).Top + 5
                End If
                If LogMove(i) = 2 Then
                    Sprite(i).Left = Sprite(i).Left - 5
                End If
                If LogMove(i) = 3 Then
                    Sprite(i).Top = Sprite(i).Top - 5
                End If
                If LogMove(i) = 4 Then
                    Sprite(i).Left = Sprite(i).Left + 5
                End If
                If LogMove(i) = 5 Then
                    Sprite(i).Top = Sprite(i).Top + 5
                End If
                If LogMove(i) = 6 Then
                    Sprite(i).Left = Sprite(i).Left - 5
                End If
                If LogMove(i) = 7 Then
                    Sprite(i).Top = Sprite(i).Top - 5
                End If
                If LogMove(i) = 8 Then
                    Sprite(i).Left = Sprite(i).Left + 5
                End If
                LogMove(i) = 0
            End If 'If bloghit = true
        Next 'q
    End If 'if sprite(i).visible = true
Next 'i

If Collision = True Then
    
    If DirLeft = True And Isaac.Left >= 0 Then
        Isaac.Left = Isaac.Left + 5
    End If
    If DirRight = True And Isaac.Left <= Me.ScaleWidth - Isaac.Width Then
        Isaac.Left = Isaac.Left - 5
    End If
    If DirUp = True And Isaac.Top >= -10 Then
        Isaac.Top = Isaac.Top + 5
    End If
    If DirDown = True And Isaac.Top + Isaac.Height <= Me.ScaleHeight Then
        Isaac.Top = Isaac.Top - 5
    End If
    
End If



    If DirLeft = True And Isaac.Left >= 0 Then
        Isaac.Num = Isaac.Num + 1
        If Isaac.Num > 9 Then Isaac.Num = 5
    End If
    If DirRight = True And Isaac.Left <= Me.ScaleWidth - 15 Then
        Isaac.Num = Isaac.Num + 1
        If Isaac.Num > 4 Then Isaac.Num = 0
    End If
    If DirUp = True Or DirDown = True Then
        Isaac.Num = Isaac.Num + 1
        If Isaac.Num > 20 Then Isaac.Num = 18
    End If
    
    Dim bFinish As Boolean
    bFinish = False
    For i = 1 To 280
        If Isaac.Left <= Plat(i).Left + 25 And Isaac.Left + Isaac.Width >= Plat(i).Left And Isaac.Top <= Plat(i).Top + 25 And Isaac.Top + Isaac.Height >= Plat(i).Top Then
            If Plat(i).Screen = 1 Then 'It is a finishing tile
                bFinish = True
                Exit For
            End If
        End If
    Next 'i
    If bFinish = True Then
        bRunning = False
        Call SetIsaacPos 'Reset Isaac's position just in case
        If hoston = True Then
            frmHost.Host.SendData "DONEARENA1" & vbCrLf
            DoEvents
        Else
            frmJoin.Client.SendData "DONEARENA1" & vbCrLf
            DoEvents
        End If
        timecount.Enabled = False
        winRace = 1
        opFinished = False
        Me.Cls
        DoEvents
        frmBattle.Show
        MsgBox "You Won The Race!"
        Unload Me
        Exit Sub
    End If
    
    

    'Draw you
    Call BitBlt(Me.hdc, Isaac.Left, Isaac.Top, picYellow(Isaac.Num).Width, picYellow(Isaac.Num).Height, picYellowM(Isaac.Num).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Isaac.Left, Isaac.Top, picYellow(Isaac.Num).Width, picYellow(Isaac.Num).Height, picYellow(Isaac.Num).hdc, 0, 0, vbSrcPaint)
    'Draw your opponent
    Call BitBlt(Me.hdc, intOpMazeX, intOpMazeY, picYellow(17).Width, picYellow(17).Height, picYellowM(15).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, intOpMazeX, intOpMazeY, picYellow(17).Width, picYellow(17).Height, picYellow(17).hdc, 0, 0, vbSrcPaint)
    
End Sub


Private Function DetectCollision() As Boolean
On Error Resume Next
Dim curCollision As Boolean
curCollision = False
For tilenum = 1 To 280
        With Plat(tilenum)
            If Isaac.Left + 4 <= .Left + 25 And Isaac.Left + Isaac.Width - 5 >= .Left And Isaac.Top + Isaac.Height - 25 <= .Top + 25 And Isaac.Top + Isaac.Height - 7 >= .Top Then
                If .Num = 41 Or .Num = 22 Or .Num = 23 Or .Num = 24 Or .Num = 25 Or .Num = 26 Or .Num = 29 Or .Num = 30 Or .Num = 54 Or .Num = 55 Or .Num = 56 Or .Num = 1 Or .Num = 4 Or .Num = 5 Or .Num = 9 Or .Num = 10 Or .Num = 11 Or .Num = 19 Or .Num = 20 Then
                    'Tile is walkable
                ElseIf (.Num = 25 Or .Num = 26) And DirDown = True Then
                    'Tile is slideable
                Else
                    curCollision = True
                    Exit For
                End If
            End If
        End With
    If tilenum <= 25 Then
        With Sprite(tilenum)
            If .Visible = True Then
                If Isaac.Left + 4 <= .Left + .Width And Isaac.Left + Isaac.Width - 5 >= .Left And Isaac.Top + Isaac.Height - 20 <= .Top + .Height And Isaac.Top + Isaac.Height - 9 >= .Top Then
                    curCollision = True
                    Exit For
                End If
            End If
        End With
    End If
Next 'tilenum
DetectCollision = curCollision
End Function

Private Sub CreateFont(Name As String, Size As Integer, Bold As Boolean, Italic As Boolean, Underline As Boolean, Color As Long)
    ' set forms font
    Me.Font.Bold = Bold
    Me.Font.Italic = Italic
    Me.Font.Name = Name
    Me.Font.Size = Size
    Me.Font.Underline = Underline
    ' set the color
    Me.ForeColor = Color
End Sub

Private Sub Form_Unload(Cancel As Integer)
bUp = False
bDown = False
bRight = False
bLeft = False
StopMidi
End Sub

Private Sub timeCount_Timer()
On Error Resume Next
If hoston = True Then
    frmHost.Host.SendData "POSMAZEX" & Isaac.Left & vbCrLf
    frmHost.Host.SendData "POSMAZEY" & Isaac.Top & vbCrLf
Else
    frmJoin.Client.SendData "POSMAZEX" & Isaac.Left & vbCrLf
    frmJoin.Client.SendData "POSMAZEY" & Isaac.Top & vbCrLf
End If

TimeLeft = TimeLeft - 1

If BattleLoaded(1) = True Then
    timecount.Enabled = False
End If

If TimeLeft = 0 And BattleLoaded(1) = False Then
    winRace = 2
    bRunning = False
    timecount.Enabled = False
    bRunning = False
    MsgBox "You have run out of time"
    DoEvents
    frmBattle.Show
    Unload Me
    Exit Sub
End If

If TimeLeft = 41 And MazeWait = True And hoston = True Then
    MazeWait = False
    TimeLeft = 45
    Call GameLoop
End If
    
End Sub
Sub DrawText()
Call CreateFont("Arial Black", 12, True, False, False, RGB(0, 0, 0))
Call TextOut(Me.hdc, 5, Me.ScaleHeight - 20, "TIME REMAINING", Len("TIME REMAINING"))
Call TextOut(Me.hdc, 200, Me.ScaleHeight - 20, CStr(TimeLeft), Len(CStr(TimeLeft)))
End Sub
Sub CheckWin()
On Error Resume Next

If opFinished = True Then 'If the opponent has won
    winRace = 2
    bRunning = False
    opFinished = False
    timecount.Enabled = False
    Call SetIsaacPos
    MsgBox "You Lost The Race!"
    DoEvents
    frmBattle.Show
    Unload Me
    Exit Sub
End If

End Sub


Sub SetIsaacPos()
    'Set Isaac's position
    Isaac.Visible = True
    Isaac.Num = 15
    Isaac.Left = 50
    Isaac.Top = -15
    IsaacX = 2
    IsaacY = 0
    Isaac.Width = picYellow(0).ScaleWidth
    Isaac.Height = picYellow(0).ScaleHeight
End Sub
