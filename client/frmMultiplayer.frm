VERSION 5.00
Begin VB.Form frmMultiplayer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts - Multiplayer"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "frmMultiplayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opPlayer 
      BackColor       =   &H00FF0000&
      Caption         =   "Player 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   78
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton opPlayer 
      BackColor       =   &H00FF0000&
      Caption         =   "Player 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   77
      Top             =   2520
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   24
      Left            =   1440
      Picture         =   "frmMultiplayer.frx":08CA
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   76
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   23
      Left            =   1080
      Picture         =   "frmMultiplayer.frx":0953
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   75
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   24
      Left            =   1440
      Picture         =   "frmMultiplayer.frx":09E4
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   74
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   23
      Left            =   1080
      Picture         =   "frmMultiplayer.frx":1061
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   73
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picIconM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   8
      Left            =   6600
      Picture         =   "frmMultiplayer.frx":15A0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   72
      Top             =   3120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIconM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   6840
      Picture         =   "frmMultiplayer.frx":18E2
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   71
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIconM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   6
      Left            =   6600
      Picture         =   "frmMultiplayer.frx":1C24
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   70
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIconM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   6840
      Picture         =   "frmMultiplayer.frx":1F66
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   69
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   8
      Left            =   6120
      Picture         =   "frmMultiplayer.frx":22A8
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   68
      Top             =   3120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   6360
      Picture         =   "frmMultiplayer.frx":2340
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   67
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   6
      Left            =   6120
      Picture         =   "frmMultiplayer.frx":26EE
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   66
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   6360
      Picture         =   "frmMultiplayer.frx":2794
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   65
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   22
      Left            =   3480
      Picture         =   "frmMultiplayer.frx":2829
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   64
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   21
      Left            =   3120
      Picture         =   "frmMultiplayer.frx":35EB
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   63
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   20
      Left            =   2760
      Picture         =   "frmMultiplayer.frx":44B9
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   62
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   22
      Left            =   1440
      Picture         =   "frmMultiplayer.frx":5387
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   61
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   21
      Left            =   1080
      Picture         =   "frmMultiplayer.frx":5998
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   60
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   20
      Left            =   1440
      Picture         =   "frmMultiplayer.frx":60AD
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   59
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   19
      Left            =   5280
      Picture         =   "frmMultiplayer.frx":67F6
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   58
      Top             =   2760
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   19
      Left            =   1080
      Picture         =   "frmMultiplayer.frx":6BA1
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   57
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   18
      Left            =   120
      Picture         =   "frmMultiplayer.frx":70C8
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   56
      Top             =   3000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   18
      Left            =   5400
      Picture         =   "frmMultiplayer.frx":8A5A
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   55
      Top             =   720
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   17
      Left            =   1440
      Picture         =   "frmMultiplayer.frx":9317
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   54
      Top             =   720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   17
      Left            =   2400
      Picture         =   "frmMultiplayer.frx":99E0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   53
      Top             =   2280
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   16
      Left            =   120
      Picture         =   "frmMultiplayer.frx":A726
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   52
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   16
      Left            =   1080
      Picture         =   "frmMultiplayer.frx":B5F4
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   51
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   15
      Left            =   6840
      Picture         =   "frmMultiplayer.frx":BD82
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   50
      Top             =   120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   15
      Left            =   960
      Picture         =   "frmMultiplayer.frx":C12B
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   49
      Top             =   2280
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   645
      Index           =   14
      Left            =   6360
      Picture         =   "frmMultiplayer.frx":C1B0
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   48
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   645
      Index           =   14
      Left            =   480
      Picture         =   "frmMultiplayer.frx":C6DF
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   47
      Top             =   2280
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   13
      Left            =   600
      Picture         =   "frmMultiplayer.frx":D899
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   46
      Top             =   1080
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   13
      Left            =   6000
      Picture         =   "frmMultiplayer.frx":E6EB
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   12
      Left            =   5400
      Picture         =   "frmMultiplayer.frx":EDEE
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   44
      Top             =   120
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   12
      Left            =   1320
      Picture         =   "frmMultiplayer.frx":F4DA
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   43
      Top             =   2280
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picIconM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   6840
      Picture         =   "frmMultiplayer.frx":108CC
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   42
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIconM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   6840
      Picture         =   "frmMultiplayer.frx":10C0E
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   41
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIconM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   6600
      Picture         =   "frmMultiplayer.frx":10F50
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   40
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIconM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   6600
      Picture         =   "frmMultiplayer.frx":11292
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   39
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIconM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   6600
      Picture         =   "frmMultiplayer.frx":115D4
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   38
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   6360
      Picture         =   "frmMultiplayer.frx":11916
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   37
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   6360
      Picture         =   "frmMultiplayer.frx":119BA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   6120
      Picture         =   "frmMultiplayer.frx":11A50
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   35
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   6120
      Picture         =   "frmMultiplayer.frx":11AD8
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   34
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   6120
      Picture         =   "frmMultiplayer.frx":11B6E
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   33
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer timeUpdate 
      Interval        =   2000
      Left            =   480
      Top             =   1800
   End
   Begin VB.TextBox txtSendMsg 
      Height          =   285
      Left            =   1680
      MaxLength       =   38
      TabIndex        =   30
      Top             =   4920
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ListBox lstGen 
      Height          =   1425
      ItemData        =   "frmMultiplayer.frx":11C17
      Left            =   1800
      List            =   "frmMultiplayer.frx":11C19
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   11
      Left            =   360
      Picture         =   "frmMultiplayer.frx":11C1B
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   10
      Left            =   480
      Picture         =   "frmMultiplayer.frx":11FD2
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   9
      Left            =   480
      Picture         =   "frmMultiplayer.frx":1238C
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   8
      Left            =   360
      Picture         =   "frmMultiplayer.frx":12575
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   7
      Left            =   6120
      Picture         =   "frmMultiplayer.frx":1274F
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   6
      Left            =   6840
      Picture         =   "frmMultiplayer.frx":128F3
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   5
      Left            =   240
      Picture         =   "frmMultiplayer.frx":12B01
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   4
      Left            =   120
      Picture         =   "frmMultiplayer.frx":12CFE
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   3
      Left            =   240
      Picture         =   "frmMultiplayer.frx":12EDA
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   2
      Left            =   360
      Picture         =   "frmMultiplayer.frx":13094
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   1
      Left            =   240
      Picture         =   "frmMultiplayer.frx":13174
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   11
      Left            =   1080
      Picture         =   "frmMultiplayer.frx":1336C
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   10
      Left            =   3960
      Picture         =   "frmMultiplayer.frx":1391B
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   9
      Left            =   4320
      Picture         =   "frmMultiplayer.frx":13EE0
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   8
      Left            =   1440
      Picture         =   "frmMultiplayer.frx":1445B
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   7
      Left            =   5040
      Picture         =   "frmMultiplayer.frx":149BB
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   6
      Left            =   4680
      Picture         =   "frmMultiplayer.frx":14F04
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   5
      Left            =   1800
      Picture         =   "frmMultiplayer.frx":1559F
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   4
      Left            =   2160
      Picture         =   "frmMultiplayer.frx":15B44
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   3
      Left            =   2520
      Picture         =   "frmMultiplayer.frx":16095
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   2
      Left            =   2880
      Picture         =   "frmMultiplayer.frx":165D1
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   1
      Left            =   3240
      Picture         =   "frmMultiplayer.frx":16B33
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   360
      Picture         =   "frmMultiplayer.frx":170DE
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   3600
      Picture         =   "frmMultiplayer.frx":1719A
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer timeMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1800
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label lblAction 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   4320
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblAction 
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   5400
      TabIndex        =   28
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblAction 
      BackStyle       =   0  'Transparent
      Caption         =   "Sell Weapon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   3480
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblAction 
      BackStyle       =   0  'Transparent
      Caption         =   "Buy Weapon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   26
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblAction 
      BackStyle       =   0  'Transparent
      Caption         =   "Examine Weapon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the Item Shop!  In here you can buy new items with coins you've earned from Single Player and Multiplayer Battles."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   600
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Shape shpText 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Shape shpSelect 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "frmMultiplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim LastTick, CurrentTick As Long
Dim TickDif As Long

Dim bWalkThroughWalls As Boolean 'Can you go through barriers?

Dim Egg18(1 To 5) As Boolean
Dim intTempEgg18 As Long

Dim bRunning As Boolean 'Is the gameloop running?

Dim bDown As Boolean
Dim bUp As Boolean
Dim bLeft As Boolean
Dim bRight As Boolean

Dim Blink As Integer

Dim LagSpam As Integer

Dim showText As Boolean

Dim curMap As String

Dim iCoinChange As Integer


Dim Sprite(1 To 25) As GameTile
Dim Plat(1 To 280) As GameTile

Private Sub Form_Activate()
bWalkThroughWalls = False

TickDif = 50
bRunning = True
Call GameLoop

End Sub

Private Sub Form_GotFocus()
bRunning = True
Call GameLoop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    bDown = True
End If
If KeyCode = vbKeyUp Then
    bUp = True
End If
If KeyCode = vbKeyLeft Then
    bLeft = True
End If
If KeyCode = vbKeyRight Then
    bRight = True
End If
If KeyCode = vbKeyReturn Then
    timeMove.Enabled = False
    txtSendMsg.Visible = True
End If
If KeyCode = vbKeyF9 Then
    bWalkThroughWalls = True
End If
If KeyCode = vbKeyR Then
    LoadMap ("Vale")
    curMap = "Vale"
    IsaacM(curIsaac).Screen = 1
    IsaacM(curIsaac).Top = 118
    IsaacM(curIsaac).Left = 193
End If
If KeyCode = vbKeyC Then
    frmChat.Show
End If
If KeyCode = vbKeyW Then
    curMap = "ItemShop"
    IsaacM(curIsaac).Screen = 10
    IsaacM(curIsaac).Left = 278
    IsaacM(curIsaac).Top = 174
    LoadMap (curMap)
End If
If KeyCode = vbKeyI Then
    curMap = "Inn"
    IsaacM(curIsaac).Screen = 6
    IsaacM(curIsaac).Left = 278
    IsaacM(curIsaac).Top = 180
    LoadMap (curMap)
End If
If KeyCode = vbKeyD Then
    curMap = "DjinnStore"
    IsaacM(curIsaac).Screen = 9
    IsaacM(curIsaac).Left = 278
    IsaacM(curIsaac).Top = 174
    LoadMap (curMap)
End If
If KeyCode = vbKeyP Then
    curMap = "PsynergyShop"
    IsaacM(curIsaac).Screen = 8
    IsaacM(curIsaac).Left = 278
    IsaacM(curIsaac).Top = 174
    LoadMap (curMap)
End If
If KeyCode = vbKeyB Then
    curMap = "BattleArena"
    IsaacM(curIsaac).Screen = 7
    IsaacM(curIsaac).Left = 278
    IsaacM(curIsaac).Top = 174
    LoadMap (curMap)
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    bDown = False
End If
If KeyCode = vbKeyUp Then
    bUp = False
End If
If KeyCode = vbKeyLeft Then
    bLeft = False
End If
If KeyCode = vbKeyRight Then
    bRight = False
End If
If KeyCode = vbKeyF9 Then
    bWalkThroughWalls = False
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

For i = 1 To 5
    Egg18(i) = False
Next 'i
intTempEgg18 = 0

'Load large picture files

frmArena.picSprite(17).Width = 242
frmArena.picSpriteM(17).Width = 242
frmArena.picSprite(17).Height = 213
frmArena.picSpriteM(17).Height = 213
frmArena.picSprite(12).Width = 216
frmArena.picSpriteM(12).Width = 216
frmArena.picSprite(12).Height = 166
frmArena.picSpriteM(12).Height = 166

Dim col As Integer
Dim row As Integer
col = 0
row = 0

Call PlayMidi("vale", True)

bRunning = False

For i = 1 To 280
    If i <= 20 Then
        intGlobalIcon(i) = 999
    End If
    Plat(i).Left = 25 * col
    Plat(i).Top = 25 * row
    Plat(i).Num = 0
    
    col = col + 1
    If col = 20 Then
        row = row + 1
        col = 0
    End If
    
    If i <= frmArena.imgTile.UBound Then
        Load frmArena.picTile(i)
        frmArena.picTile(i).Picture = frmArena.imgTile(i).Picture
        frmArena.picTile(i).Width = 25
        frmArenapic.Tile(i).Height = 25
        frmArenapic.Tile(i).Left = 25 * (i - 1)
        frmArenapic.Tile(i).Visible = False
    End If
    If i <= 100 Then
        Sprite(i).Visible = False
    End If
    If i <= 20 Then
        IsaacM(i).Visible = False
        Load lblmsg(i)
        lblmsg(i).Caption = ""
        lblmsg(i).Visible = False
        lblmsg(i).BackColor = lblmsg(0).BackColor
        lblmsg(i).ForeColor = lblmsg(0).ForeColor
        lblmsg(i).Width = lblmsg(0).Width
        lblmsg(i).Height = lblmsg(0).Height
        lblmsg(i).AutoSize = True
        lblmsg(i).Font = lblmsg(0).Font
    End If
    
Next 'i



strMap = "Vale"

curMap = "Vale"

opFinished = False

Call LoadMap(strMap)

timeMove.Enabled = True
LagSpam = 0

showText = False

Blink = 0


frmArena.picSprite(17).Picture = LoadPicture(App.Path & "\HouseInside.gif")
frmArena.picSpriteM(17).Picture = LoadPicture(App.Path & "\HouseInsideM.gif")
'242x213
frmArena.picSprite(17).Width = 242
frmArena.picSpriteM(17).Width = 242
frmArena.picSprite(17).Height = 213
frmArena.picSpriteM(17).Height = 213

frmArena.picSprite(12).Picture = LoadPicture(App.Path & "\Lake01.gif")
frmArena.picSpriteM(12).Picture = LoadPicture(App.Path & "\Lake01M.gif")
'216x166
frmArena.picSprite(12).Width = 216
frmArena.picSpriteM(12).Width = 216
frmArena.picSprite(12).Height = 166
frmArena.picSpriteM(12).Height = 166


bMazeFirstLoad = True


End Sub

Sub LoadMap(ByVal strSave As String)
On Error Resume Next

If Egg18(3) = False And strSave = "ValeSouth" Or strSave = "ValeEast" Then
    intTempEgg18 = intTempEgg18 + 1
End If
If intTempEgg18 = 20 Then
    If Egg18(1) = True And Egg18(2) = True Then
        Egg18(3) = True
    End If
    Beep
    intTempEgg18 = 0
End If

strSave = App.Path & "\" & strSave & ".dat"
For i = 1 To 280
    Dim intCur As Integer
    intCur = CInt(GetFromIni("GEN", "T" & i, strSave))
    Plat(i).Num = intCur
    Plat(i).Width = 25
    Plat(i).Height = 25
    
    Dim strTileLink As String
    strTileLink = GetFromIni("GEN", "G" & i, strSave)
    
    If strTileLink = "T" Then
        Plat(i).Link = GetFromIni("GEN", "GL" & i, strSave)
    Else
        Plat(i).Link = ""
    End If
    
    If i <= 25 Then
        Sprite(i).Num = CInt(GetFromIni("GEN", "STYPE" & i, strSave))
        Sprite(i).Height = frmArena.picSprite(Sprite(i).Num).ScaleHeight
        Sprite(i).Width = frmArena.picSprite(Sprite(i).Num).ScaleWidth
        Sprite(i).Left = CInt(GetFromIni("GEN", "SLEFT" & i, strSave))
        Sprite(i).Top = CInt(GetFromIni("GEN", "STOP" & i, strSave))
        If Sprite(i).Num <> 999 Then
            Sprite(i).Visible = True
        Else
            Sprite(i).Visible = False
        End If
        Sprite(i).IconTime = 0
    End If
        
Next 'i
End Sub

Private Sub Form_LostFocus()
bRunning = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print X & "      " & Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
bRunning = False
StopMidi
End Sub

Private Sub lblAction_Click(Index As Integer)
On Error Resume Next
Dim iCoins As Long
iCoins = CLng(strCoins)
If Index = 0 Then
    If lblAction(0).Caption = "Find Games" Then
        lstGen.Clear
        frmChat.Chat.SendData "GETGAMELIST" & vbCrLf
    End If
    If lblAction(0).Caption = "Examine Character" Then
        If lstGen.Text = "Menardi" Then
            If Egg18(1) = True And Egg18(2) = False And Egg18(3) = False And Egg18(4) = False And Egg18(5) = False Then
                Egg18(2) = True
                Beep
            End If
        End If
        lblText.Caption = GetCharDesc(lstGen.Text)
    End If
    If lblAction(0).Caption = "Examine Psynergy" Then
        Dim intRealPsy As Integer
        For i = 1 To 30
            If strPsyName(i) = lstGen.Text Then
                intRealPsy = i
            End If
        Next 'i
        
        lblText.Caption = strPsyName(intRealPsy) & ": " & strPsyDesc(intRealPsy) & ".  Type: " & strPsyType(intRealPsy) & ". Djinni Required: " & strPsyDjinn(intRealPsy) & ". PP Required: " & strPsyPP(intRealPsy)
    End If
    If lblAction(0).Caption = "Examine Weapon" Then
        lblText.Caption = strItemName(lstGen.ListIndex + 1) & ": " & strItemDesc(lstGen.ListIndex + 1) & ".  Price: " & strItemCoins(lstGen.ListIndex + 1) & " coins.  Attack power: " & strItemDamage(lstGen.ListIndex + 1) & "."
    End If
    If lblAction(0).Caption = "Examine Djinni" Then
        Dim intRealDjinn As Integer
        For i = 1 To 20
            If lstGen.Text = Djinn(i).Name Then
                intRealDjinn = i
            End If
        Next 'i
        If Djinn(intRealDjinn).State = 0 Then
            lblText.Caption = strDjinnName(intRealDjinn) & ": Attack Type: " & strDjinnType(intRealDjinn) & ".  SET"
        ElseIf Djinn(intRealDjinn).State = 1 Then
            lblText.Caption = strDjinnName(intRealDjinn) & ": Attack Type: " & strDjinnType(intRealDjinn) & ".  STANDBY"
        Else
            lblText.Caption = strDjinnName(intRealDjinn) & ": Attack Type: " & strDjinnType(intRealDjinn) & ".  REST"
        End If
    End If
End If
If Index = 1 Then
    If lblAction(1).Caption = "Create Game" Then
        frmHost.Show

        frmHost.lblmsg.Caption = "No users connected."
        frmHost.Host.Close
        frmHost.cmdStart.Enabled = False
        frmHost.txtGameName.Enabled = True
        frmHost.cmdCreate.Enabled = True
        frmHost.chkGen(0).Enabled = False
        frmHost.chkGen(1).Enabled = False
        frmHost.cmdBoot.Enabled = False
        frmHost.cmdArena.Enabled = False
        frmHost.txtmsg.Enabled = False
        frmHost.cmdSend.Enabled = False
    End If
    If lblAction(1).Caption = "Switch Character" Then
        If iCoins < iCoinChange Then
            lblText.Caption = "Not enough coins!  You currently have " & iCoins & " coins and to change you need " & iCoinChange & " coins."
        Else
            iCoins = iCoins - iCoinChange
            strCoins = CStr(iCoins)
            
            Dim iCurCust As Long
            iCurCust = FindWhichCharacter(lstGen.Text)
            
            lblText.Caption = "Character switched to " & lstGen.Text & ".  You currently have " & iCoins & " coins remaining.  Please restart the game for the effect to take place."
            If opPlayer(0).Value = True Then
                If iCurCust = 999 Then
                    frmChat.Chat.SendData "SWITCHCHAR" & lstGen.Text & vbCrLf
                Else
                    frmChat.Chat.SendData "SWITCHCUSTCHAR" & iCurCust & vbCrLf
                End If
            Else
                If iCurCust = 999 Then
                    frmChat.Chat.SendData "2SWITCHCHAR" & lstGen.Text & vbCrLf
                Else
                    frmChat.Chat.SendData "2SWITCHCUSTCHAR" & iCurCust & vbCrLf
                End If
            End If
            frmChat.Chat.SendData "KILLCOINS" & iCoinChange & vbCrLf
        End If
    End If
    If lblAction(1).Caption = "Buy Weapon" Then
        If lstGen.Text <> "" Then
            If iCoins >= CLng(strItemCoins(lstGen.ListIndex + 1)) Then
                If lstGen.Text = "Hunter's Sword" Then
                    Dim bTempEgg As Boolean
                    bTempEgg = True
                    For i = 1 To 5
                         If Egg18(i) = True Then
                            bTempEgg = False
                        End If
                    Next 'i
                    If bTempEgg = True Then
                        Egg18(1) = True
                        Beep
                    End If
                End If
                iCoins = iCoins - CInt(strItemCoins(lstGen.ListIndex + 1))
                strCoins = CStr(iCoins)
                frmUser2.User.SendData "NEWITEMUSER" & strMyUserName & vbCrLf
                If opPlayer(0).Value = True Then
                    frmUser2.User.SendData "NEWITEMCHAR" & CStr(1) & vbCrLf
                Else
                    frmUser2.User.SendData "NEWITEMCHAR" & CStr(2) & vbCrLf
                End If
                frmUser2.User.SendData "NEWITEMCOINS" & CStr(iCoins) & vbCrLf
                frmUser2.User.SendData "NEWITEMNAME" & strItemName(lstGen.ListIndex + 1) & vbCrLf
                
                
                
                lblText.Caption = "Bought item!"
            Else
                lblText.Caption = "Not enough coins!"
            End If
        End If
    End If
    If lblAction(1).Caption = "Toggle Djinni State" Then
        For i = 1 To 20
            If lstGen.Text = Djinn(i).Name Then
                intRealDjinn = i
            End If
        Next 'i
        If Djinn(intRealDjinn).State = 0 Then
            Djinn(intRealDjinn).State = 1
            MsgBox "Djinn unset"
            If Egg18(1) = True And Egg18(2) = True And Egg18(3) = True Then
                Egg18(4) = True
                Beep
            End If
        Else
            MsgBox "Djinn set"
            Djinn(intRealDjinn).State = 0
        End If
    End If
End If
If Index = 2 Then
    If lblAction(2).Caption = "Join Game" Then
        frmJoin.Client.Close
        frmChat.Chat.SendData "JOINGAME" & lstGen.Text & vbCrLf
        frmJoin.txtip.Enabled = True
        frmJoin.cmdListen.Enabled = True
        frmJoin.cmdSend.Enabled = False
        frmJoin.txtmsg.Enabled = False
        frmJoin.lblmsg.Caption = "Not connected."

        frmJoin.chkReady.Enabled = False
    End If
End If
    
If Index = 3 Then
    timeMove.Enabled = True
    lblAction(0).Visible = False
    lblAction(1).Visible = False
    lblAction(2).Visible = False
    lblAction(3).Visible = False
    lblAction(4).Visible = False
    shpSelect.Visible = False
    shpText.Visible = False
    lblText.Visible = False
    lstGen.Visible = False
    lstGen.Clear
    Blink = 100
End If

If Index = 4 Then
    If lblAction(4).Caption = "Manually Find Game" Then
        frmJoin.Show
        frmJoin.txtip.Visible = True
        frmJoin.cmdListen.Visible = True
        strJoinIP = ""
    End If
End If

End Sub

Sub DisplayText()
On Error Resume Next
            Select Case curMap
                Case "Inn"
                    lblAction(0).Caption = "Examine Character"
                    lblAction(1).Caption = "Switch Character"
                    lblAction(2).Caption = ""
                    lblAction(3).Caption = "Quit"
                    lblAction(4).Caption = ""
                    
                    Dim strCharCoins As String
                    iCoinChange = 10 * CInt(strLvl)
                    strCharCoins = CStr(iCoinChange)
                    
                    lblText.Caption = "Welcome to the inn, you can switch characters here.  It costs " & strCharCoins & " coins to switch.  You currently have " & strCoins & " coins.  Your current character is " & strChar(1)
                    
                    For q = 0 To frmUser2.cmbCharPic.ListCount - 1
                        lstGen.AddItem frmUser2.cmbCharPic.List(q)
                    Next 'q
                Case "ItemShop"
                    lstGen.Clear
                    
                    For q = 1 To 30
                        If strItemName(q) <> "" Then
                            lstGen.AddItem strItemName(q)
                        End If
                    Next 'q
                    
                    lblAction(0).Caption = "Examine Weapon"
                    lblAction(1).Caption = "Buy Weapon"
                    lblAction(2).Caption = ""
                    lblAction(3).Caption = "Quit"
                    lblAction(4).Caption = ""
                    lblText.Caption = "Welcome to the Item Shop.  Here you can buy weapons using coins gained from battle.  Weapons will automatically be equipped.  Your current weapon is " & lstGen.List(CInt(strWeapon(1))) & ".  You currently have " & strCoins & " coins."
                Case "DjinnStore"
                    lblAction(0).Caption = "Examine Djinni"
                    lblAction(1).Caption = "Toggle Djinni State"
                    lblAction(2).Caption = ""
                    lblAction(3).Caption = "Quit"
                    lblAction(4).Caption = ""
                    'iCoinChange = CInt(strLvl) * 100
                    lblText.Caption = "Welcome to the Djinn Farm.  In here you can toggle the state of your Djinn from Set to to Standby."
                    For q = 1 To 10
                        If strDjinnName(q) <> "" Then
                            lstGen.AddItem strDjinnName(q)
                        End If
                    Next 'q
                Case "PsynergyShop"
                    lblAction(0).Caption = "Examine Psynergy"
                    lblAction(1).Caption = ""
                    lblAction(2).Caption = ""
                    lblAction(3).Caption = "Quit"
                    lblAction(4).Caption = ""
                    lblText.Caption = "Welcome to the Psynergy Shop.  In this store you can examine the current Psynergy that you have."
                    For q = 1 To 30
                        If strPsyName(q) <> "" Then
                            lstGen.AddItem strPsyName(q)
                        End If
                    Next 'q
                Case "BattleArena"
                    lblAction(0).Caption = "Find Games"
                    lblAction(1).Caption = "Create Game"
                    lblAction(2).Caption = "Join Game"
                    lblAction(3).Caption = "Quit"
                    lblAction(4).Caption = "Manually Find Game"
                    frmChat.Chat.SendData "GETGAMELIST" & vbCrLf
                    lblText.Caption = "Welcome to the Battle Arena.  From here you can create or join a game.  To find the current games avalible, press the Find Games button."
                Case "House"
                    lblAction(0).Caption = ""
                    lblAction(1).Caption = ""
                    lblAction(2).Caption = ""
                    lblAction(3).Caption = "Quit"
                    lblAction(4).Caption = ""
                    lblText.Caption = "Welcome to your house.  In here you can send stats to the Single Player version of the game and get coins gained from Single Player (note: You can't do this yet.)."
                Case "SecretHouse"
                    lblAction(0).Caption = ""
                    lblAction(1).Caption = ""
                    lblAction(2).Caption = ""
                    lblAction(3).Caption = "Quit"
                    lblAction(4).Caption = ""
                    lblText.Caption = "You have found the secret house!  Congratulations, you are officially a cool guy (or girl :).  Easter Egg #16"
                    Call Encode("16", "EGG16", "EGGL16", App.Path & "\settings.ini")
                Case "TopSecretZone"
                    lblAction(0).Caption = ""
                    lblAction(1).Caption = ""
                    lblAction(2).Caption = ""
                    lblAction(3).Caption = "Quit"
                    lblAction(4).Caption = ""
                    lblText.Caption = "You have found the mother of all easter eggs, #18.  My congratulations to you, you've earned it!  Easter Egg #18"
                    Call Encode("18", "EGG18", "EGGL18", App.Path & "\settings.ini")
            End Select
End Sub

Private Sub timeUpdate_Timer()
On Error Resume Next

    frmChat.Chat.SendData "SCREEN" & IsaacM(curIsaac).Screen & vbCrLf
    DoEvents
    frmChat.Chat.SendData "ISAACX" & IsaacM(curIsaac).Left & vbCrLf
    DoEvents
    frmChat.Chat.SendData "ISAACY" & IsaacM(curIsaac).Top & vbCrLf
    DoEvents

End Sub

Private Sub txtSendMsg_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And txtSendMsg.Text <> "" Then
    If txtSendMsg.Text = "/go" Then
        curMap = InputBox("TopSecretZone,SecretHouse,SecretZone", "Don't leave empty, may crash!", "SecretZone")
        Blink = 50
        IsaacM(curIsaac).Left = Me.ScaleWidth / 2
        IsaacM(curIsaac).Top = Me.ScaleHeight / 2
        Call LoadMap(curMap)
    End If
    If txtSendMsg.Text = "/custom" Then
        curMap = InputBox("Kenny,Goten,Duel,Great Healer,SnoMan,Trunks,Yoshi,Vegito,Gardevoir,Nathan Graves,Sabre Master,Evil Cheese,Jason,Roy,Vegeta,Cloud,King of Dark, Great Volkan,Celes,Link,Ray,MetaKnight,Kenshin,RoboBug,Ganondorf,Crono,Barbarian,Eclipse,Master Chief,Laguna.", "Don't leave empty, may crash!", "Crono")
            frmChat.Chat.SendData "SWITCHCUSTCHAR" & FindWhichCharacter(curMap) & vbCrLf
        txtSendMsg.Text = ""
        txtSendMsg.Visible = False
        timeMove.Enabled = True
    End If
    If Left$(txtSendMsg.Text, 2) = "/!" Then
        frmChat.Chat.SendData "IN!" & vbCrLf
    ElseIf Left$(txtSendMsg.Text, 2) = "/?" Then
        frmChat.Chat.SendData "IN?" & vbCrLf
    ElseIf Left$(txtSendMsg.Text, 4) = "/..." Then
        frmChat.Chat.SendData "INDOT" & vbCrLf
    ElseIf Left$(txtSendMsg.Text, 3) = "/:)" Then
        frmChat.Chat.SendData "INSMILE" & vbCrLf
    ElseIf Left$(txtSendMsg.Text, 3) = "/:(" Then
        frmChat.Chat.SendData "INFROWN" & vbCrLf
    ElseIf Left$(txtSendMsg.Text, 5) = "/love" Then
        frmChat.Chat.SendData "INLOVE" & vbCrLf
    ElseIf Left$(txtSendMsg.Text, 3) = "/:[" Then
        frmChat.Chat.SendData "INANGRY" & vbCrLf
    ElseIf Left$(txtSendMsg.Text, 5) = "/idea" Then
        frmChat.Chat.SendData "INIDEA" & vbCrLf
    ElseIf Left$(txtSendMsg.Text, 2) = "/+" Then
        frmChat.Chat.SendData "INCLOUD" & vbCrLf
    ElseIf Left$(txtSendMsg.Text, 4) = "/egg" Then
        MsgBox "You have found yet another easter egg.  Go pat yourself on the back.", vbInformation, "Easter Egg #30"
        Call Encode("30", "EGG30", "EGGL30", App.Path & "\settings.ini")
    Else
        If txtSendMsg.Text = "kenny" Then
            If Egg18(1) = True And Egg18(2) = True And Egg18(3) = True And Egg18(4) = True Then
                curMap = "TopSecretZone"
                Blink = 50
                IsaacM(curIsaac).Left = Me.ScaleWidth / 2
                IsaacM(curIsaac).Top = Me.ScaleHeight / 2
                Call LoadMap(curMap)
            Else
                frmChat.Chat.SendData "INGAMECHAT" & strMyUserName & ": " & txtSendMsg.Text & vbCrLf
            End If
        Else
            frmChat.Chat.SendData "INGAMECHAT" & strMyUserName & ": " & txtSendMsg.Text & vbCrLf
        End If
    End If
    
    txtSendMsg.Text = ""
    txtSendMsg.Visible = False
    timeMove.Enabled = True
    
End If
If KeyCode = vbKeyEscape Then
    txtSendMsg.Text = ""
    txtSendMsg.Visible = False
    timeMove.Enabled = True
End If

End Sub
Function GetCharDesc(ByVal strCurChar As String) As String
Dim intCurCust As Long
intCurCust = FindWhichCharacter(strCurChar)

If strCurChar = "Isaac" Then
GetCharDesc = "Isaac: An all around character with good Attack."

ElseIf strCurChar = "Garret" Then
GetCharDesc = "Garret: Good HP, bad PP, average Attack, poor Luck."

ElseIf strCurChar = "Jenna" Then
GetCharDesc = "Jenna: Good PP, Average HP, Average Attack, Psynergy suffers 1/5 damage loss, average luck."

ElseIf strCurChar = "Ivan" Then
GetCharDesc = "Ivan: Good PP, Bad HP, Bad Attack, Gains 1/4 Psynergy Damage Bonus, good luck."

ElseIf strCurChar = "Mia" Then
GetCharDesc = "Mia: Great PP, Bad Attack, Average HP, Suffers 1/4 Psynergy Damage Loss, great luck."

ElseIf strCurChar = "Sheba" Then
GetCharDesc = "Sheba: Average HP, Average PP, Bad Attack. Gains 1/3 Psynergy Damage Bonus.  Has average luck."

ElseIf strCurChar = "Felix" Then
GetCharDesc = "Felix: Great Attack, Average HP, Bad PP, Average Luck."

ElseIf strCurChar = "Alex" Then
GetCharDesc = "Alex: Good HP, Good Attack, Bad PP.  Has poor luck."

ElseIf strCurChar = "Saturos" Then
GetCharDesc = "Saturos: Great HP, Great Attack, Bad PP.  Psynergy Cost Doubled.  Bad luck."

ElseIf strCurChar = "Menardi" Then
GetCharDesc = "Menardi: Good HP, Average PP, Average Attack, Poor Luck."

ElseIf strCurChar = "Kraden" Then
GetCharDesc = "Kraden: Exceptional PP, Bad HP, Bad Attack.  Gains 1/3 Psynergy Damage Bonus, Bad Luck."

ElseIf strCurChar = "Caption Contest Character" Then
GetCharDesc = "Lizard Man has extremely high HP and AP.  However, he has very low Psynergy."

ElseIf strCurChar = "Guard" Then
GetCharDesc = "Guard: No Psynergy, Weak Attack, Weak Defense, Low Luck.  EXCEPTIONAL leveling-up stats."

ElseIf strCurChar = "Gladiator" Then
GetCharDesc = "Gladiator: Self proclaimed dominator of low-level play.  Great initial Attack, Defense, Luck.  No Psynergy."

ElseIf strCurChar = "Piers" Then
    GetCharDesc = "Piers: A powerful sea pirate that specializes in his high power and resist."

ElseIf strCurChar = "Kenny" Then
    GetCharDesc = "Kenny: A resurected zombie with a firey urge to eat brains.  Incredible statistics.  Weak to water."
ElseIf strCurChar = "KOS" Then
    GetCharDesc = "Absolute K oS: A well rounded character like Isaac with good strength, defense but no Psynergy."
ElseIf strCurChar = "Cloud" Then
    GetCharDesc = "Cloud Strife: Thanks to his FF7 skills, he excells in magic.  He is not so good in Attack and Resist."

ElseIf strCurChar = "Purple Piers" Then
    GetCharDesc = "Purple Piers: Piers, but with gnarly purple hair."

ElseIf strCurChar = "Agiato" Then
    GetCharDesc = "Agiato: Master of the Dark Flame.  Strong Psynergy, Average HP, Low Luck, Low Resist, High Power."

ElseIf strCurChar = "Karst" Then
    GetCharDesc = "Karst: The younger, darker sister of Menardi.  Average Psynergy, Average HP, Low Luck, Average Resist, High Power."

ElseIf strCurChar = "The Wise One" Then
    GetCharDesc = "The Wise One: Guardian of the Sol Sanctum, he is the ultimate elemental master.  Exceptional Psynergy, Power and Resist.  No AP, No Defense."
ElseIf strCurChar = "Young Isaac" Then
    GetCharDesc = "Young Isaac: A younger, more energetic Isaac.  He has less HP and PP but greater everything else."
ElseIf strCurChar = "Young Garet" Then
    GetCharDesc = "Young Garet: Very high AP and Luck but low Defense and Resist."
Else
    If intCurCust <> 999 Then
        GetCharDesc = CustomChar(intCurCust).Name & ": " & CustomChar(intCurCust).Description
    End If
End If

End Function

Sub CharMove()
On Error Resume Next
If bDown = True And IsaacM(curIsaac).Top <= Me.ScaleHeight - 15 Then
    IsaacM(curIsaac).Top = IsaacM(curIsaac).Top + 3
End If
If bUp = True And IsaacM(curIsaac).Top >= -15 Then
    IsaacM(curIsaac).Top = IsaacM(curIsaac).Top - 3
End If
If bLeft = True And IsaacM(curIsaac).Left >= 0 Then
    IsaacM(curIsaac).Left = IsaacM(curIsaac).Left - 3
End If
If bRight = True And IsaacM(curIsaac).Left <= Me.ScaleWidth - 10 Then
    IsaacM(curIsaac).Left = IsaacM(curIsaac).Left + 3
End If


Dim Collide As Boolean

Collide = False

For i = 1 To 280
    If IsaacM(curIsaac).Left + 5 <= Plat(i).Left + Plat(i).Width And IsaacM(curIsaac).Left + IsaacM(curIsaac).Width - 5 >= Plat(i).Left And IsaacM(curIsaac).Top + Abs(25 - IsaacM(curIsaac).Height) + 4 <= Plat(i).Top + Plat(i).Height And (IsaacM(curIsaac).Top + Abs(25 - IsaacM(curIsaac).Height)) + 21 >= Plat(i).Top Then
        With Plat(i)
            If bWalkThroughWalls = False And .Num <> 54 And .Num <> 55 And .Num <> 56 And .Num <> 1 And .Num <> 4 And .Num <> 5 And .Num <> 19 And .Num <> 20 And .Num <> 9 And .Num <> 10 And .Num <> 11 And .Num <> 27 And .Num <> 28 And .Num <> 21 And .Num <> 22 And .Num <> 23 And .Num <> 29 And .Num <> 30 And .Num <> 41 Then
                Collide = True
            End If
            
            If Blink = 0 And .Link <> "" Then
                Select Case .Link
                
                Case "ValeWest"
                    IsaacM(curIsaac).Screen = 2
                    IsaacM(curIsaac).Left = Me.ScaleWidth - IsaacM(curIsaac).Width
                Case "ValeNorth"
                    IsaacM(curIsaac).Screen = 3
                    IsaacM(curIsaac).Top = Me.ScaleHeight - IsaacM(curIsaac).Height
                Case "ValeEast"
                If curMap = "ValeSouth" Then
                    IsaacM(curIsaac).Screen = 4
                    IsaacM(curIsaac).Left = 0
                    IsaacM(curIsaac).Top = 208
                    curMap = "ValeEast"
                End If
                If curMap = "ItemShop" Then
                    IsaacM(curIsaac).Screen = 4
                    IsaacM(curIsaac).Left = 147
                    IsaacM(curIsaac).Top = 94
                    curMap = "ValeEast"
                End If
                Case "ValeSouth"
                    If curMap = "ValeEast" Then
                    IsaacM(curIsaac).Left = Me.ScaleWidth - 20
                    IsaacM(curIsaac).Top = 208
                    End If
                    If curMap = "Vale" Then
                    IsaacM(curIsaac).Top = 0
                    IsaacM(curIsaac).Left = 250
                    End If
                    If curMap = "PsynergyShop" Then
                    IsaacM(curIsaac).Top = 157
                    IsaacM(curIsaac).Left = 102
                    End If
                    If curMap = "Djinn Store" Then
                    IsaacM(curIsaac).Top = 280
                    IsaacM(curIsaac).Left = 328
                    End If
                    
                    IsaacM(curIsaac).Screen = 5
                    curMap = "ValeSouth"
                    
                Case "Vale"
                
                    IsaacM(curIsaac).Screen = 1
                    
                    
                    
                    If curMap = "Inn" Then
                        IsaacM(curIsaac).Left = 193
                        IsaacM(curIsaac).Top = 113
                    End If
                    If curMap = "BattleArena" Then
                        IsaacM(curIsaac).Left = 68
                        IsaacM(curIsaac).Top = 286
                    End If
                    If curMap = "ValeSouth" Then
                        IsaacM(curIsaac).Top = Me.ScaleHeight - IsaacM(curIsaac).Height
                    End If
                    If curMap = "SecretZone" Then
                        IsaacM(curIsaac).Top = 286
                        IsaacM(curIsaac).Left = 68
                    End If
                    
                    
                    curMap = "Vale"
                    
                
                Case "Inn"
                    curMap = "Inn"
                    IsaacM(curIsaac).Screen = 6
                    IsaacM(curIsaac).Left = 278
                    IsaacM(curIsaac).Top = 180
                Case "BattleArena"
                    curMap = "BattleArena"
                    IsaacM(curIsaac).Screen = 7
                    IsaacM(curIsaac).Left = 278
                    IsaacM(curIsaac).Top = 174
                Case "PsynergyShop"
                    curMap = "PsynergyShop"
                    IsaacM(curIsaac).Screen = 8
                    IsaacM(curIsaac).Left = 278
                    IsaacM(curIsaac).Top = 174
                Case "House"
                    curMap = "DjinnStore"
                    IsaacM(curIsaac).Screen = 9
                    IsaacM(curIsaac).Left = 278
                    IsaacM(curIsaac).Top = 174
                Case "ItemShop"
                    curMap = "ItemShop"
                    IsaacM(curIsaac).Screen = 10
                    IsaacM(curIsaac).Left = 278
                    IsaacM(curIsaac).Top = 174
                Case "ValeNE"
                    curMap = "ValeNE"
                    IsaacM(curIsaac).Screen = 11
                    IsaacM(curIsaac).Left = 278
                    IsaacM(curIsaac).Top = 174
                Case "SecretZone"
                    curMap = "SecretZone"
                    IsaacM(curIsaac).Screen = 49
                    IsaacM(curIsaac).Left = Plat(209).Left
                    IsaacM(curIsaac).Top = Plat(209).Top
                Case "SecretHouse"
                    curMap = "SecretHouse"
                    IsaacM(curIsaac).Screen = 50
                    IsaacM(curIsaac).Left = Plat(249).Left
                    IsaacM(curIsaac).Top = Plat(249).Top
                End Select
                
            LoadMap (curMap)
            
            Blink = 100
            
            Exit Sub
            
            End If
        End With
    End If
Next 'i

For i = 1 To 25
    If IsaacM(curIsaac).Left <= Sprite(i).Left + Sprite(i).Width And IsaacM(curIsaac).Left + IsaacM(curIsaac).Width >= Sprite(i).Left And IsaacM(curIsaac).Top + Abs(25 - IsaacM(curIsaac).Height) <= Sprite(i).Top + Sprite(i).Height And (IsaacM(curIsaac).Top + Abs(25 - IsaacM(curIsaac).Height)) + 25 >= Sprite(i).Top Then
        If Blink = 0 And Sprite(i).Num = 16 Then
            lblAction(0).Visible = True
            lblAction(1).Visible = True
            lblAction(2).Visible = True
            lblAction(3).Visible = True
            lblAction(4).Visible = True
            shpSelect.Visible = True
            shpText.Visible = True
            lblText.Visible = True
            lstGen.Visible = True
            lstGen.Clear
            showText = True
            
            Call DisplayText
            
            Collide = True
        End If
    End If
                    
        
Next 'i

If Collide = True Then
    If bDown = True Then
        IsaacM(curIsaac).Top = IsaacM(curIsaac).Top - 3
    End If
    If bUp = True Then
        IsaacM(curIsaac).Top = IsaacM(curIsaac).Top + 3
    End If
    If bLeft = True Then
        IsaacM(curIsaac).Left = IsaacM(curIsaac).Left + 3
    End If
    If bRight = True Then
        IsaacM(curIsaac).Left = IsaacM(curIsaac).Left - 3
    End If
End If

If IsaacM(curIsaac).Left + IsaacM(curIsaac).Width >= Me.ScaleWidth And IsaacM(curIsaac).Top + IsaacM(curIsaac).Height >= Me.ScaleHeight Then
    curMap = "SecretZone"
    IsaacM(curIsaac).Screen = 49
    IsaacM(curIsaac).Top = Plat(249).Top
    IsaacM(curIsaac).Left = Plat(249).Left
    Blink = 100
    LoadMap ("SecretZone")
End If



Me.Cls

For i = 1 To 280
    Call BitBlt(Me.hdc, Plat(i).Left, Plat(i).Top, 25, 25, frmArena.picTile(Plat(i).Num).hdc, 0, 0, vbSrcCopy)
Next '
For i = 1 To 25
    If i <= 25 Then
        If Sprite(i).Visible = True Then
            Call BitBlt(Me.hdc, Sprite(i).Left, Sprite(i).Top, Sprite(i).Width, Sprite(i).Height, frmArena.picSpriteM(Sprite(i).Num).hdc, 0, 0, vbSrcAnd)
            Call BitBlt(Me.hdc, Sprite(i).Left, Sprite(i).Top, Sprite(i).Width, Sprite(i).Height, frmArena.picSprite(Sprite(i).Num).hdc, 0, 0, vbSrcPaint)
        End If
    End If

Next 'i

i = curIsaac

For i = 1 To 20
    If IsaacM(i).Visible = True And IsaacM(i).Screen = IsaacM(curIsaac).Screen Then
        If IsaacM(i).CustomCharacter = False Then
        Call BitBlt(Me.hdc, IsaacM(i).Left, IsaacM(i).Top, picChar(IsaacM(i).Num).ScaleWidth, picChar(IsaacM(i).Num).ScaleHeight, picCharM(IsaacM(i).Num).hdc, 0, 0, vbSrcAnd)
        Call BitBlt(Me.hdc, IsaacM(i).Left, IsaacM(i).Top, picChar(IsaacM(i).Num).ScaleWidth, picChar(IsaacM(i).Num).ScaleHeight, picChar(IsaacM(i).Num).hdc, 0, 0, vbSrcPaint)
        Else
        Call BitBlt(Me.hdc, IsaacM(i).Left, IsaacM(i).Top, frmChat.picCustChar(IsaacM(i).Num).ScaleWidth, frmChat.picCustChar(IsaacM(i).Num).ScaleHeight, frmChat.picCustCharM(IsaacM(i).Num).hdc, 0, 0, vbSrcAnd)
        Call BitBlt(Me.hdc, IsaacM(i).Left, IsaacM(i).Top, frmChat.picCustChar(IsaacM(i).Num).ScaleWidth, frmChat.picCustChar(IsaacM(i).Num).ScaleHeight, frmChat.picCustChar(IsaacM(i).Num).hdc, 0, 0, vbSrcPaint)
        End If
    End If
Next 'i



If showText = True Then
    showText = False
    timeMove.Enabled = False
    bDown = False
    bUp = False
    bLeft = False
    bRight = False
End If

If Blink > 0 Then
    Blink = Blink - 1
End If

For i = 1 To 20
    If DispText(i) <> 0 Then
        DispText(i) = DispText(i) - 1
        Call DrawText(i)
    ElseIf DispText(i) = 0 Then
        lblmsg(i).Visible = False
    End If
    If intGlobalIcon(i) <> 999 Then
        Sprite(i).IconTime = 100
        Sprite(i).IconType = intGlobalIcon(i)
        intGlobalIcon(i) = 999
    End If
    If Sprite(i).IconTime > 0 And IsaacM(i).Screen = IsaacM(curIsaac).Screen Then
        Sprite(i).IconTime = Sprite(i).IconTime - 1
        Call BitBlt(Me.hdc, IsaacM(i).Left + 5, (IsaacM(i).Top - picIcon(Sprite(i).IconType).ScaleHeight), picIcon(Sprite(i).IconType).ScaleWidth, picIcon(Sprite(i).IconType).ScaleHeight, picIconM(Sprite(i).IconType).hdc, 0, 0, vbSrcAnd)
        Call BitBlt(Me.hdc, IsaacM(i).Left + 5, (IsaacM(i).Top - picIcon(Sprite(i).IconType).ScaleHeight), picIcon(Sprite(i).IconType).ScaleWidth, picIcon(Sprite(i).IconType).ScaleHeight, picIcon(Sprite(i).IconType).hdc, 0, 0, vbSrcPaint)
    End If
Next 'i


DoEvents

End Sub
Sub GameLoop()
On Error Resume Next
Do While bRunning = True And Me.txtSendMsg.Visible = False
    CurrentTick = GetTickCount()
    If CurrentTick - LastTick > TickDif And lstGen.Visible = False Then
        LastTick = CurrentTick
        Call CharMove
    End If
    DoEvents
Loop
End Sub
Sub DrawText(ByVal i As Integer)
On Error Resume Next
    Me.Font.Bold = False
    Me.Font.Italic = False
    Me.Font.Name = "Verdana"
    Me.Font.Size = 10
    Me.Font.Underline = False
    Me.ForeColor = RGB(0, 0, 0)
    
    Call TextOut(Me.hdc, IsaacM(i).Left + 25, IsaacM(i).Top, lblmsg(i).Caption, Len(lblmsg(i).Caption))
    Me.ForeColor = RGB(255, 255, 255)

    Call TextOut(Me.hdc, IsaacM(i).Left + 27, IsaacM(i).Top + 2, lblmsg(i).Caption, Len(lblmsg(i).Caption))


End Sub
