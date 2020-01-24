VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts - Chat"
   ClientHeight    =   6990
   ClientLeft      =   2340
   ClientTop       =   2370
   ClientWidth     =   7425
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmChat.frx":08CA
   Picture         =   "frmChat.frx":1194
   ScaleHeight     =   466
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   495
   Begin VB.ListBox lstScores 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2985
      ItemData        =   "frmChat.frx":1A5E
      Left            =   5160
      List            =   "frmChat.frx":1A60
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.VScrollBar vUser 
      Height          =   2895
      LargeChange     =   8
      Left            =   7200
      Max             =   8
      TabIndex        =   96
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer timeFlash 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   0
   End
   Begin MSComDlg.CommonDialog comDiag 
      Left            =   2640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.maz"
      Filter          =   "WOTA Maze Files (*.maz) |*.maz|"
      InitDir         =   "app.path"
   End
   Begin VB.Frame framProfile 
      BackColor       =   &H00808080&
      Caption         =   "Player's Profile"
      Height          =   2295
      Left            =   120
      TabIndex        =   53
      Top             =   4440
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton cmdPAvatar 
         Caption         =   "&Change"
         Height          =   255
         Left            =   4680
         TabIndex        =   108
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtPURL 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   93
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtPICQ 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   73
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdPHide 
         Caption         =   "&Hide"
         Height          =   255
         Left            =   6120
         TabIndex        =   71
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdPSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5160
         TabIndex        =   70
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtPOther 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   75
         TabIndex        =   69
         Top             =   2280
         Width           =   6495
      End
      Begin VB.TextBox txtPEmail 
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   68
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtPMSN 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   67
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtPAIM 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   66
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtPLocation 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   65
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtPSex 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   64
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtPAge 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   63
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtPName 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   62
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avatar:"
         Height          =   195
         Index           =   15
         Left            =   3000
         TabIndex        =   94
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URL:"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   92
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image imgProfChar 
         Height          =   225
         Left            =   3600
         Picture         =   "frmChat.frx":1A62
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblPKarma 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   3600
         TabIndex        =   86
         Top             =   2880
         Width           =   120
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KARMA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   13
         Left            =   2760
         TabIndex        =   85
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label lblPCharacter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Isaac"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1440
         TabIndex        =   82
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHARACTER:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   81
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblPCoins 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   3720
         TabIndex        =   80
         Top             =   2640
         Width           =   330
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COINS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   11
         Left            =   3000
         TabIndex        =   79
         Top             =   2640
         Width           =   645
      End
      Begin VB.Label lblPRating 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2400
         TabIndex        =   78
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RATING:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   10
         Left            =   1560
         TabIndex        =   77
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label lblPLoss 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1200
         TabIndex        =   76
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label lblPWins 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   840
         TabIndex        =   75
         Top             =   2640
         Width           =   210
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STATS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   74
         Top             =   2640
         Width           =   675
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICQ:"
         Height          =   195
         Index           =   8
         Left            =   2400
         TabIndex        =   72
         Top             =   840
         Width           =   315
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   61
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   195
         Index           =   6
         Left            =   4680
         TabIndex        =   60
         Top             =   840
         Width           =   465
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MSN:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   59
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AIM:"
         Height          =   195
         Index           =   4
         Left            =   5280
         TabIndex        =   58
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location:"
         Height          =   195
         Index           =   3
         Left            =   2400
         TabIndex        =   57
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   56
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   55
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblProfGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame framScrambler 
      BackColor       =   &H00404080&
      Caption         =   "Scrambler Bot"
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CheckBox chkQuickClues 
         BackColor       =   &H00404080&
         Caption         =   "Quicker Clues"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3120
         TabIndex        =   43
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdScoreList 
         Caption         =   "View Scores"
         Height          =   255
         Left            =   4560
         TabIndex        =   42
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Timer timeHighScore 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2520
         Top             =   600
      End
      Begin VB.HScrollBar hSWinners 
         Height          =   255
         LargeChange     =   3
         Left            =   3480
         Max             =   3
         Min             =   1
         TabIndex        =   38
         Top             =   1440
         Value           =   1
         Width           =   975
      End
      Begin VB.Timer timeAnswer 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   1800
         Top             =   600
      End
      Begin VB.CheckBox chkAnswers 
         BackColor       =   &H00404080&
         Caption         =   "Auto Scroll Clues After 30 Seconds"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   35
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmdSRules 
         Caption         =   "Show &Rules"
         Height          =   255
         Left            =   4560
         TabIndex        =   34
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdSHighScores 
         Caption         =   "Show &High Scores"
         Height          =   255
         Left            =   4560
         TabIndex        =   33
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Timer timeSWait 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   2160
         Top             =   600
      End
      Begin VB.CommandButton cmdScramble 
         Caption         =   "Scramble"
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtSCategory 
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   30
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdSHelp 
         Caption         =   "&Help"
         Height          =   315
         Left            =   6120
         TabIndex        =   29
         Top             =   1680
         Width           =   975
      End
      Begin VB.HScrollBar hSPoints 
         Height          =   255
         LargeChange     =   10
         Left            =   1920
         Max             =   100
         Min             =   1
         TabIndex        =   28
         Top             =   1440
         Value           =   1
         Width           =   1575
      End
      Begin VB.CommandButton cmdSClear 
         Caption         =   "&Clear List"
         Height          =   255
         Left            =   6000
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdSNewGame 
         Caption         =   "&Reset Scores"
         Height          =   255
         Left            =   4560
         TabIndex        =   24
         Top             =   1800
         Width           =   1455
      End
      Begin VB.FileListBox filSList 
         Height          =   870
         Left            =   1800
         Pattern         =   "*.srm"
         TabIndex        =   22
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdSLoad 
         Caption         =   "&Add List"
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox chkList 
         BackColor       =   &H00404080&
         Caption         =   "Scramble From List"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtSScrambled 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtSOrgWord 
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.ListBox lstSWords 
         Height          =   840
         ItemData        =   "frmChat.frx":1E62
         Left            =   4320
         List            =   "frmChat.frx":1E64
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdSExit 
         Caption         =   "E&xit"
         Height          =   315
         Left            =   6120
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblSWinners 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4200
         TabIndex        =   37
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblScrambler 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Winners:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   3480
         TabIndex        =   36
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lblScrambler 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   675
      End
      Begin VB.Label lblSPoints 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3000
         TabIndex        =   27
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lblScrambler 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Points Worth:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   26
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblScrambler 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List To Use:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   1800
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblScrambler 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scrambled Phrase:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblScrambler 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phrase To Scramble:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblScrambler 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Word List:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4320
         TabIndex        =   15
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.OptionButton opChat 
      BackColor       =   &H00000000&
      Caption         =   "Messages"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   90
      Top             =   3675
      Width           =   1095
   End
   Begin VB.ListBox lstIgnore 
      Height          =   255
      ItemData        =   "frmChat.frx":1E66
      Left            =   4200
      List            =   "frmChat.frx":1E68
      TabIndex        =   89
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmd30Freeze 
      BackColor       =   &H00E0E0E0&
      Caption         =   "60 Second Freeze"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdPause 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pa&use Chat"
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdProfile 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Profile"
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3960
      Width           =   735
   End
   Begin VB.PictureBox picCustChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   690
      Index           =   0
      Left            =   1920
      Picture         =   "frmChat.frx":1E6A
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   51
      Top             =   -240
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picCustCharM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   690
      Index           =   0
      Left            =   2280
      Picture         =   "frmChat.frx":2071
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   50
      Top             =   -240
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.OptionButton opChat 
      BackColor       =   &H00000000&
      Caption         =   "Games"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   47
      Top             =   3675
      Width           =   855
   End
   Begin VB.OptionButton opChat 
      BackColor       =   &H00000000&
      Caption         =   "View All"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   44
      Top             =   3675
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.Timer timeFreeze 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2520
      Top             =   3120
   End
   Begin VB.Timer timeType 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1440
      Top             =   3120
   End
   Begin VB.Timer timeHelp 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   1800
      Top             =   3120
   End
   Begin VB.Timer timeAway 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2160
      Top             =   3120
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Help With Chat Commands (Double Click to Skip)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Picture         =   "frmChat.frx":2278
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdTown 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Launch Town Window"
      Height          =   975
      Left            =   120
      MaskColor       =   &H000080FF&
      Picture         =   "frmChat.frx":2582
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Report &User"
      Height          =   255
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Timer timeSpam 
      Interval        =   3500
      Left            =   3480
      Top             =   0
   End
   Begin VB.CommandButton cmdMsg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Msg"
      Height          =   255
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   735
   End
   Begin VB.ListBox lstCheck 
      Height          =   3375
      ItemData        =   "frmChat.frx":2AA0
      Left            =   7200
      List            =   "frmChat.frx":2AA2
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Send"
      Height          =   255
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   735
   End
   Begin VB.ListBox lstUsers 
      BackColor       =   &H00FFFFFF&
      Height          =   2985
      ItemData        =   "frmChat.frx":2AA4
      Left            =   7200
      List            =   "frmChat.frx":2AA6
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtMsg 
      BackColor       =   &H00886000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      MaxLength       =   265
      TabIndex        =   0
      Top             =   3960
      Width           =   4935
   End
   Begin MSWinsockLib.Winsock Chat 
      Left            =   7080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9887
   End
   Begin VB.CommandButton cmdIgnore 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ignore"
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame framAvatar 
      BackColor       =   &H00808080&
      Caption         =   "Select Your Avatar"
      Height          =   3015
      Left            =   120
      TabIndex        =   109
      Top             =   600
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSaveAvatar 
         Caption         =   "&Save"
         Height          =   255
         Left            =   3960
         TabIndex        =   113
         Top             =   2640
         Width           =   855
      End
      Begin VB.FileListBox filAvatar 
         Height          =   2625
         Left            =   240
         Pattern         =   "*.gif"
         TabIndex        =   110
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   18
         Left            =   4080
         Picture         =   "frmChat.frx":2AA8
         Stretch         =   -1  'True
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   17
         Left            =   3600
         Picture         =   "frmChat.frx":2DD7
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   16
         Left            =   3120
         Picture         =   "frmChat.frx":3175
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   15
         Left            =   2640
         Picture         =   "frmChat.frx":3517
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   14
         Left            =   2160
         Picture         =   "frmChat.frx":38B6
         Top             =   1440
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   13
         Left            =   3600
         Picture         =   "frmChat.frx":3C4F
         Top             =   2280
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   255
         Index           =   12
         Left            =   3120
         Picture         =   "frmChat.frx":41A9
         Top             =   2280
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   11
         Left            =   2640
         Picture         =   "frmChat.frx":45DD
         Top             =   2280
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   10
         Left            =   2160
         Picture         =   "frmChat.frx":4B15
         Top             =   2280
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   9
         Left            =   4080
         Picture         =   "frmChat.frx":4FCB
         Top             =   1920
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   8
         Left            =   3600
         Picture         =   "frmChat.frx":527F
         Top             =   1920
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   7
         Left            =   3120
         Picture         =   "frmChat.frx":572B
         Top             =   1920
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   6
         Left            =   2640
         Picture         =   "frmChat.frx":5BD1
         Top             =   1920
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   5
         Left            =   2160
         Picture         =   "frmChat.frx":5CC3
         Top             =   1920
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   4
         Left            =   4080
         Picture         =   "frmChat.frx":60F6
         Top             =   1680
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   3
         Left            =   3600
         Picture         =   "frmChat.frx":6471
         Top             =   1680
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   2
         Left            =   3120
         Picture         =   "frmChat.frx":6509
         Top             =   1680
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   1
         Left            =   2640
         Picture         =   "frmChat.frx":6881
         Top             =   1680
         Width           =   375
      End
      Begin VB.Image imgSAvatar 
         Height          =   225
         Index           =   0
         Left            =   2160
         Picture         =   "frmChat.frx":6BFA
         Top             =   1680
         Width           =   375
      End
      Begin VB.Image imgAvatarPreview 
         Height          =   225
         Left            =   2880
         Picture         =   "frmChat.frx":6FF5
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Preview:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   112
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblGen 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Please Note: Selecting a non-official custom avatar will only display on players with that file's computers"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   0
         Left            =   2160
         TabIndex        =   111
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.TextBox txtMessage 
      ForeColor       =   &H0000C000&
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   91
      Top             =   600
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtTalk 
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   84
      Top             =   600
      Visible         =   0   'False
      Width           =   4935
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   8937472
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":73F5
   End
   Begin RichTextLib.RichTextBox txtGame 
      Height          =   3015
      Left            =   120
      TabIndex        =   48
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   8937472
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":748D
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   107
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   106
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   105
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   104
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   5640
      TabIndex        =   103
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5640
      TabIndex        =   102
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   101
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   100
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   99
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   98
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   97
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   11
      Left            =   5160
      Picture         =   "frmChat.frx":7525
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   10
      Left            =   5160
      Picture         =   "frmChat.frx":7928
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   9
      Left            =   5160
      Picture         =   "frmChat.frx":7D28
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   8
      Left            =   5160
      Picture         =   "frmChat.frx":8137
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   7
      Left            =   5160
      Picture         =   "frmChat.frx":84AF
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   6
      Left            =   5160
      Picture         =   "frmChat.frx":88B1
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   5
      Left            =   5160
      Picture         =   "frmChat.frx":8D3C
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   4
      Left            =   5160
      Picture         =   "frmChat.frx":91CD
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   3
      Left            =   5160
      Picture         =   "frmChat.frx":964C
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   2
      Left            =   5160
      Picture         =   "frmChat.frx":9ADA
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   1
      Left            =   5160
      Picture         =   "frmChat.frx":9E53
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   0
      Left            =   5160
      Picture         =   "frmChat.frx":A25B
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   95
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblChatFilter 
      BackColor       =   &H00000000&
      Caption         =   "Select Chat Filter:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblUsersOn 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   46
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lblUsersOnline 
      BackColor       =   &H00000000&
      Caption         =   "Users Online:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   45
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H00000040&
      Caption         =   "Lost?  Click Here To View The Game's Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   720
      MouseIcon       =   "frmChat.frx":A6E7
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label lblHide 
      BackColor       =   &H00000000&
      Caption         =   "Hide Banner"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5760
      TabIndex        =   39
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   1200
      Left            =   120
      MouseIcon       =   "frmChat.frx":AFB1
      MousePointer    =   99  'Custom
      Picture         =   "frmChat.frx":B87B
      Top             =   5520
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.Label lblCrotch 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   5880
      Width           =   495
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   344
      X2              =   480
      Y1              =   35
      Y2              =   35
   End
   Begin VB.Label lblUsers 
      BackStyle       =   0  'Transparent
      Caption         =   "USERS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5160
      TabIndex        =   10
      Top             =   -120
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   8
      X2              =   336
      Y1              =   35
      Y2              =   35
   End
   Begin VB.Label lblBanner 
      BackStyle       =   0  'Transparent
      Caption         =   "CHAT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   -120
      Width           =   1815
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intFilterEgg As Long
Dim iAnswers As Integer 'Seconds before answers start
Dim bScrambler As Boolean 'Is scrambler on or not?
Dim strMacro(0 To 1) As String
Dim SPlayers(1 To 30) As Scrambler
Dim HostScramble As Boolean 'Am I the host?
Dim CurSItem As Integer 'Current list item used
Dim TotalWinners As Integer 'Total people who have won
Dim intSHS As Integer 'Scrambler high score
Dim intHighScore As Integer 'Current high score
Dim intHighestPlayer As Integer 'Player who commands the highest score
Dim bScramBusy As Boolean
Dim intFreeze As Long
Dim strAwayMessage As String

Dim strCustomAvatar As String

Dim MinDisplayPlayer As Long
Dim MaxDisplayPlayer As Long

Dim strLastMessage As String


Dim intFlash As Long
Dim bFocus As Boolean


Dim strLastToSendMessage As String 'Last person that sent you a message

Dim bScramblerGame As Boolean
Dim iKarmaSpam As Long
Dim iMaxKarma As Long
Dim intFreezeMax As Long

Dim strProfileAvatar As String
Dim intDifficulty As Integer 'Difficulty of the typing game
Dim TypeOrDie As Boolean 'Determines whether or not if the timer is activated again and the word isn't typed the player loses or not
Dim strWord As String 'Current word that needs to be typed
Dim intTypeWord As Integer 'Current display on the typing test
Dim intHelp As Integer 'The current text to display in the help timer
Dim bAway As Boolean
Dim intAway As Integer
Dim intTempIcon As Integer
Dim Away As Boolean
Dim intChatNum As Integer
Dim CurMoveIsaac As Integer
Dim ChatUsers(1 To 20) As ChatUser
Dim Admin As Boolean
Dim Moderator As Boolean
Dim SuperModerator As Boolean
Dim iSpam As Integer

Private Sub Chat_Connect()
On Error Resume Next
txtChat.Text = txtChat.Text & vbNewLine & "Connected to server!"
Chat.SendData "CHATNAME" & frmUser2.txtUserName.Text & vbCrLf
DoEvents
If intKarma > 150 Then
    Dim strSplit
    strSplit = Split(frmIntro.txtWOTAVet.Text, "[me]", -1, vbTextCompare)
    Dim strVetMSG As String
    For i = 0 To UBound(strSplit)
        If i = 0 Then
            strVetMSG = strVetMSG & strSplit(i)
        Else
            strVetMSG = strVetMSG & strMyUserName & strSplit(i)
        End If
    Next 'i
    Chat.SendData "GOLDTXT" & strVetMSG & vbCrLf
    frmIntro.txtWOTAVet.Enabled = True
ElseIf intKarma > 45 Then
    Chat.SendData "MODTXT" & "High Karma User " & strMyUserName & " has signed on." & vbCrLf
    DoEvents
End If
End Sub

Private Sub Chat_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
'Comment out for Ladder Tournament

Dim strdatao As String
Chat.GetData strdatao
Dim strTime As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
strData = Split(strdatao, vbCrLf, -1, vbTextCompare)
For i = 0 To UBound(strData)
    If Left$(strData(i), 5) = "WHERE" Then
        Dim strloc As String, intx As Integer, inty As Integer, intCurWhere As Integer
        intCurWhere = CInt(Mid$(strData(i), 6, Len(strData(i))))
        
        txtloc = Switch(IsaacM(intCurWhere).Screen = 1, "Vale", IsaacM(intCurWhere).Screen = 2, "ValeWest", IsaacM(intCurWhere).Screen = 3, "ValeNorth", IsaacM(intCurWhere).Screen = 4, "ValeEast", IsaacM(intCurWhere).Screen = 5, "ValeSouth", IsaacM(intCurWhere).Screen = 6, "Inn", IsaacM(intCurWhere).Screen = 7, "BattleArena", IsaacM(intCurWhere).Screen = 8, "PsynergyShop", IsaacM(intCurWhere).Screen = 9, "Djinn Shop", IsaacM(intCurWhere).Screen = 10, "ItemShop", IsaacM(intCurWhere).Screen = 49, "Cool Guy Zone", IsaacM(intCurWhere).Screen = 50, "Cool Guy Zone")
        intx = IsaacM(intCurWhere).Left
        inty = IsaacM(intCurWhere).Top
        strChatTxt = "Player is located at: " & txtloc & " (" & intx & "/" & inty & ")"
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strChatTxt)
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &HC000&
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelText = vbNewLine & strChatTxt
        
    End If
    
    If Left$(strData(i), 7) = "MODWARN" Then
        Call AddAllChat("You have been warned by a moderator.  Your Karma has gone down by 1 and if you continue to violate the Terms of Service you may be kicked or banned.")
    End If
    If Left$(strData(i), 9) = "MODPRAISE" Then
        Call AddAllChat("You have been praised by a moderator, good job.  Your Karma has increased by 1.")
    End If
    
    If Left$(strData(i), 9) = "GETJOINIP" Then
        Dim strGetJoinIP As String
        strGetJoinIP = Mid$(strData(i), 10, Len(strData(i)))
        strJoinIP = strGetJoinIP
        frmJoin.Show
        frmJoin.txtip.Text = strGetJoinIP
    End If
    
    If Left$(strData(i), 6) = "REALIP" Then
        strRealIP = Mid$(strData(i), 7, Len(strData(i)))
    End If
    
    If Left$(strData(i), 5) = "PUSER" Then
        framProfile.Caption = Mid$(strData(i), 6, Len(strData(i)))
    End If
    
    If Left$(strData(i), 5) = "PNAME" Then
        txtPName.Text = Mid$(strData(i), 6, Len(strData(i)))
    End If
    If Left$(strData(i), 4) = "PAGE" Then
        txtPAge.Text = Mid$(strData(i), 5, Len(strData(i)))
    End If
    If Left$(strData(i), 4) = "PSEX" Then
        txtPSex.Text = Mid$(strData(i), 5, Len(strData(i)))
    End If
    If Left$(strData(i), 9) = "PLOCATION" Then
        txtPLocation.Text = Mid$(strData(i), 10, Len(strData(i)))
    End If
    If Left$(strData(i), 4) = "PAIM" Then
        txtPAIM.Text = Mid$(strData(i), 5, Len(strData(i)))
    End If
    If Left$(strData(i), 4) = "PMSN" Then
        txtPMSN.Text = Mid$(strData(i), 5, Len(strData(i)))
    End If
    If Left$(strData(i), 6) = "PEMAIL" Then
        txtPEmail.Text = Mid$(strData(i), 7, Len(strData(i)))
    End If
    If Left$(strData(i), 6) = "POTHER" Then
        txtPOther.Text = Mid$(strData(i), 7, Len(strData(i)))
    End If
    If Left$(strData(i), 5) = "PWINS" Then
        lblPWins.Caption = Mid$(strData(i), 6, Len(strData(i)))
        lblPWins.Caption = lblPWins.Caption & "/"
    End If
    If Left$(strData(i), 5) = "PLOSS" Then
        lblPLoss.Caption = Mid$(strData(i), 6, Len(strData(i)))
    End If
    If Left$(strData(i), 7) = "PRATING" Then
        lblPRating.Caption = Mid$(strData(i), 8, Len(strData(i)))
    End If
    If Left$(strData(i), 5) = "PCOINS" Then
        lblPCoins.Caption = Mid$(strData(i), 6, Len(strData(i)))
    End If
    If Left$(strData(i), 7) = "PAVATAR" Then
        imgProfChar.Picture = LoadPicture(App.Path & "\icons\" & Mid$(strData(i), 8, Len(strData(i))) & ".gif")
        strProfileAvatar = Mid$(strData(i), 8, Len(strData(i)))
    End If
    If Left$(strData(i), 4) = "PURL" Then
        txtPURL.Text = Mid$(strData(i), 5, Len(strData(i)))
    End If
    
    If Left$(strData(i), 10) = "PCHARACTER" Then
        lblPCharacter.Caption = Mid$(strData(i), 11, Len(strData(i)))
        Dim intProfPic As Long
        intProfPic = FindWhichCharacter(lblPCharacter.Caption)
        'If intProfPic = 999 Then
        '    imgProfChar.Visible = False
        'Else
        '    imgProfChar.Visible = True
        '    imgProfChar.Picture = LoadPicture(App.Path & "\files\" & CustomChar(intProfPic).Picture & ".gif")
        'End If
    End If
    If Left$(strData(i), 4) = "PICQ" Then
        txtPICQ.Text = Mid$(strData(i), 5, Len(strData(i)))
    End If
    If Left$(strData(i), 6) = "PKARMA" Then
        lblPKarma.Caption = Mid$(strData(i), 7, Len(strData(i)))
    End If
    If Left$(strData(i), 5) = "GETAP" Then
        Chat.SendData "MODTXT" & strMyUserName & "'s AP is " & AP(1) & vbCrLf
    End If
    If Left$(strData(i), 5) = "GETPP" Then
        Chat.SendData "MODTXT" & strMyUserName & "'s PP is " & PP(1) & vbCrLf
    End If
    'If Left$(strdata(i), 9) = "GETMAZEXY" Then
    '    Chat.SendData "MODTXT" & strMyUserName & "'s Current maze position is " & frmArena.Isaac.Left & " / " & frmArena.Isaac.Top & vbCrLf
    'End If
    If Left$(strData(i), 9) = "CLOSEGAME" Then
        Unload frmBattle
        MsgBox "Your game has been closed by an administrator."
    End If
    
    If Left$(strData(i), 5) = "MODON" Then
        Dim strModOn As String
        strModOn = Mid$(strData(i), 6, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = strModOn Then
                Users(q).Moderator = True
            End If
        Next 'q
        Call UpdateUserDisplay
    End If
    If Left$(strData(i), 6) = "MODOFF" Then
        strModOn = Mid$(strData(i), 7, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = strModOn Then
                Users(q).Moderator = False
                Users(q).Admin = False
            End If
        Next 'q
        Call UpdateUserDisplay
    End If
    If Left$(strData(i), 7) = "ADMINON" Then
        strModOn = Mid$(strData(i), 8, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = strModOn Then
                Users(q).Admin = True
            End If
        Next 'q
        Call UpdateUserDisplay
    End If
    If Left$(strData(i), 8) = "ADMINOFF" Then
        strModOn = Mid$(strData(i), 9, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = strModOn Then
                Users(q).Admin = False
            End If
        Next 'q
        Call UpdateUserDisplay
    End If
    
    
    

 '    If Left$(strdata(i), 9) = "WHEREBACK" Then
'        'Dim strChatMsg As String
'        strChatMsg = Mid(strdata(i), 10, Len(strdata(i)))
'        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
'        frmChat.txtChat.SelLength = Len(strChatMsg)
'        frmChat.txtChat.SelBold = True
'        frmChat.txtChat.SelColor = &H80FF&
'        frmChat.txtChat.SelFontName = "Verdana Ref"
'        frmChat.txtChat.SelText = vbNewLine & strChatMsg
'    End If


    If Left$(strData(i), 7) = "GOLDTXT" Then
        strChatTxt = Mid$(strData(i), 8, Len(strData(i)))
        If strChatTxt <> "" Then
            frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
            frmChat.txtChat.SelLength = Len(strChatTxt)
            frmChat.txtChat.SelBold = True
            frmChat.txtChat.SelColor = &HFF80FF
            frmChat.txtChat.SelFontName = "Arial"
            frmChat.txtChat.SelText = vbNewLine & strChatTxt
        End If
    End If
    
    'If Left$(strdata(i), 7) = "TALKTXT" Then
    '    strChatTxt = Mid(strdata(i), 8, Len(strdata(i)))

        
'        strNoImp = Split(strChatTxt, ":", 2, vbTextCompare)
'
'
'        If strChatTxt <> "" Then
'        frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
'        frmChat.txtTalk.SelLength = Len(strNoImp(0))
'        frmChat.txtTalk.SelBold = True
'        frmChat.txtTalk.SelColor = &HB35937
'        frmChat.txtTalk.SelFontName = "Arial"
'        frmChat.txtTalk.SelText = vbNewLine & strNoImp(0) & ":"
'
'
'        For q = 1 To UBound(strNoImp)
'
'            frmChat.txtTalk.SelStart = Len(frmtalk.txtChat.Text)
'            frmChat.txtTalk.SelLength = Len(strNoImp(q))
'            frmChat.txtTalk.SelBold = False
'            If Mid$(strdata(i), 8, Len(strMyUserName)) = strMyUserName Then
'                frmChat.txtTalk.SelColor = &H6FCCAB
'            Else
'                frmChat.txtTalk.SelColor = &HF
'            End If
'            frmChat.txtTalk.SelFontName = "MS Sans Serif"
'            frmChat.txtTalk.SelText = strNoImp(q)

'        Next 'q
        
        
'        End If
        
        
'    End If
    If Left$(strData(i), 7) = "GAMETXT" Then
        '&H78BCF3
        strChatTxt = Mid(strData(i), 8, Len(strData(i)))

        strNoImp = Split(strChatTxt, ":", 2, vbTextCompare)
        
        'Dim bIgnore As Boolean
        bIgnore = CheckIgnore(CStr(strNoImp(0))) 'Check for ignoring
        
        If bIgnore = False And strChatTxt <> "" Then
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strNoImp(0))
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &HB35937
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelText = vbNewLine & strNoImp(0) & ":"
        
        
        
        For q = 1 To UBound(strNoImp)
        
            frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
            frmChat.txtChat.SelLength = Len(strNoImp(q))
            frmChat.txtChat.SelBold = False
            If Mid$(strData(i), 8, Len(strMyUserName)) = strMyUserName Then
                frmChat.txtChat.SelColor = &H78BCF3
            Else
                frmChat.txtChat.SelColor = &HF
            End If
            frmChat.txtChat.SelFontName = "Arial"
            frmChat.txtChat.SelText = strNoImp(q)

        Next 'q
        
        
        
        If InStr(1, strChatTxt, strMacro(0), vbTextCompare) > 0 Or InStr(1, strChatTxt, strMyUserName, vbTextCompare) > 0 Then
            frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
            frmChat.txtGame.SelLength = Len(strChatTxt)
            frmChat.txtGame.SelBold = False
            If Mid$(strData(i), 8, Len(strMyUserName)) = strMyUserName Then
                frmChat.txtGame.SelColor = &H808000
            Else
                frmChat.txtGame.SelColor = &HF
            End If
            frmChat.txtGame.SelFontName = "MS Sans Serif"
            frmChat.txtGame.SelText = vbNewLine & strChatTxt
        End If
        

        
        
        End If
        
        If strChatTxt <> "" And HostScramble = True And cmdScramble.Caption = "Stop" Then 'If I'm hosting and the game is playing
            'Dim strSText
            Call DetectScramble(strChatTxt)
        End If
        
    End If

    
    If Left$(strData(i), 8) = "ADMINTXT" Then
        'If chatLoaded = False Then
        'frmChat.Show
        'End If
        strChatTxt = Mid(strData(i), 9, Len(strData(i)))
        Dim strNewLine
        strNewLine = Split(strChatTxt, "%n", -1, vbTextCompare)
        
        If bLogChat = True Then
            Call WriteIni("CHAT", strTime, CStr(strChatTxt), App.Path & "\" & strServerDate & ".ini")
        End If
        
        For q = 0 To UBound(strNewLine)
            If strNewLine(q) <> "" Then
            frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
            frmChat.txtChat.SelLength = Len(strNewLine(q))
            frmChat.txtChat.SelBold = True
            frmChat.txtChat.SelColor = &H80&
            '&H00C0C0FF&
            frmChat.txtChat.SelItalic = False
            frmChat.txtChat.SelFontName = "Arial"
            frmChat.txtChat.SelText = vbNewLine & strNewLine(q)
            
            frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
            frmChat.txtGame.SelLength = Len(strNewLine(q))
            frmChat.txtGame.SelBold = True
            frmChat.txtGame.SelColor = &H80&
            frmChat.txtGame.SelFontName = "Verdana Ref"
            frmChat.txtGame.SelText = vbNewLine & strNewLine(q)
            
            'frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
            'frmChat.txtTalk.SelLength = Len(strNewLine(q))
            'frmChat.txtTalk.SelBold = True
            'frmChat.txtTalk.SelColor = &HFF
            'frmChat.txtTalk.SelFontName = "Verdana Ref"
            'frmChat.txtTalk.SelText = vbNewLine & strNewLine(q)
            End If
        Next 'q
    End If
    'Comment out for ladder tournament
    If Left$(strData(i), 6) = "MODTXT" Then
        'If chatLoaded = False Then
        'frmChat.Show
        'End If
        strChatTxt = Mid(strData(i), 7, Len(strData(i)))
        If bLogChat = True Then
            Call WriteIni("CHAT", strTime, CStr(strChatTxt), App.Path & "\" & strServerDate & ".ini")
        End If
        If strChatTxt <> "" Then
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strChatTxt)
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &HFFFF&
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelItalic = False
        frmChat.txtChat.SelText = vbNewLine & strChatTxt
        
        frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
        frmChat.txtGame.SelLength = Len(strChatTxt)
        frmChat.txtGame.SelBold = True
        frmChat.txtGame.SelColor = &H80&
        frmChat.txtGame.SelFontName = "Arial"
        frmChat.txtGame.SelText = vbNewLine & strChatTxt
        
        'frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
        'frmChat.txtTalk.SelLength = Len(strChatTxt)
        'frmChat.txtTalk.SelBold = True
        'frmChat.txtTalk.SelColor = &HFF
        'frmChat.txtTalk.SelFontName = "Verdana Ref"
        'frmChat.txtTalk.SelText = vbNewLine & strChatTxt
        End If
    End If
    If Left$(strData(i), 7) = "CHATMSG" Then
        'Dim strChatMsg As String
        strChatMsg = Mid(strData(i), 8, Len(strData(i)))
        strChatTxt = strChatMsg
        
        If bLogChat = True Then
            Call WriteIni("MSG", strTime, CStr(strChatTxt), App.Path & "\" & strServerDate & ".ini")
        End If
        
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strChatMsg)
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &HC000&
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelItalic = False
        frmChat.txtChat.SelText = vbNewLine & strChatMsg
        
        frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
        frmChat.txtGame.SelLength = Len(strChatTxt)
        frmChat.txtGame.SelBold = True
        frmChat.txtGame.SelColor = &HC000&
        frmChat.txtGame.SelFontName = "Arial"
        frmChat.txtGame.SelText = vbNewLine & strChatTxt
        
        txtMessage.Text = txtMessage.Text & vbNewLine & strChatTxt
        
        'frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
        'frmChat.txtTalk.SelLength = Len(strChatTxt)
        'frmChat.txtTalk.SelBold = True
        'frmChat.txtTalk.SelColor = &HC000&
        'frmChat.txtTalk.SelFontName = "Verdana Ref"
        'frmChat.txtTalk.SelText = vbNewLine & strChatTxt
    End If
    If Left$(strData(i), 8) = "ANNOUNCE" Then
        Dim strAnnounce As String
        strAnnounce = Mid$(strData(i), 9, Len(strData(i)))
        MsgBox "An Announcement From The Server: " & vbNewLine & strAnnounce, vbCritical, "Server Announcement"
    End If
    If Left$(strData(i), 7) = "SMODTXT" Then
        'If chatLoaded = False Then
        'frmChat.Show
        'End If
        strChatTxt = Mid(strData(i), 8, Len(strData(i)))
        If bLogChat = True Then
            Call WriteIni("CHAT", strTime, CStr(strChatTxt), App.Path & "\" & strServerDate & ".ini")
        End If
        If strChatTxt <> "" Then
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strChatTxt)
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &HFFFF&
        frmChat.txtChat.SelItalic = True
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelText = vbNewLine & strChatTxt
        
        frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
        frmChat.txtGame.SelLength = Len(strChatTxt)
        frmChat.txtGame.SelBold = True
        frmChat.txtChat.SelItalic = True
        frmChat.txtGame.SelColor = &HFF0000
        frmChat.txtGame.SelFontName = "Arial"
        frmChat.txtGame.SelText = vbNewLine & strChatTxt
        
        'frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
        'frmChat.txtTalk.SelLength = Len(strChatTxt)
        'frmChat.txtTalk.SelBold = True
        'frmChat.txtTalk.SelColor = &HFF
        'frmChat.txtTalk.SelFontName = "Verdana Ref"
        'frmChat.txtTalk.SelText = vbNewLine & strChatTxt
        End If
    End If
    If Left$(strData(i), 10) = "CHATFREEZE" And SuperModerator = False And Admin = False Then
        intFreezeMax = 3
        intFreeze = 0
        timeFreeze.Enabled = False
        DoEvents
        timeFreeze.Enabled = True
        frmMultiplayer.Show
        frmChat.Enabled = False
        frmIntro.Enabled = False
        frmMultiplayer.Enabled = False
        strChatMsg = "You have been frozen for three minutes by a Moderator.  Until then, you will be unable to chat."
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strChatMsg)
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &HFF0000
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelText = vbNewLine & strChatMsg
    End If
    If Left$(strData(i), 8) = "60FREEZE" Then
        If Admin = False And Moderator = False And SuperModerator = False Then
            intFreezeMax = 1
            intFreeze = 0
            timeFreeze.Enabled = False
            DoEvents
            timeFreeze.Enabled = True
            frmMultiplayer.Show
            frmChat.Enabled = False
            frmIntro.Enabled = False
            frmMultiplayer.Enabled = False
            strChatMsg = "You have been frozen for one minute by a War of the Adepts Veteran.  Until then, you will be unable to chat."
            frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
            frmChat.txtChat.SelLength = Len(strChatMsg)
            frmChat.txtChat.SelBold = True
            frmChat.txtChat.SelColor = &HFF0000
            frmChat.txtChat.SelFontName = "Arial"
            frmChat.txtChat.SelText = vbNewLine & strChatMsg
        End If
    End If
    'Comment out for ladder tournament
    If Left$(strData(i), 5) = "METXT" Then
        'Dim strChatMsg As String
        strChatMsg = Mid(strData(i), 6, Len(strData(i)))
        strChatTxt = strChatMsg
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strChatMsg)
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &H80FF&
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelText = vbNewLine & strChatMsg
        
        'frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
        'frmChat.txtTalk.SelLength = Len(strChatTxt)
        'frmChat.txtTalk.SelBold = True
        'frmChat.txtTalk.SelColor = &H80FF&
        'frmChat.txtTalk.SelFontName = "Verdana Ref"
        'frmChat.txtTalk.SelText = vbNewLine & strChatTxt
    End If
    If Left$(strData(i), 7) = "AWAYTXT" Then
        'Dim strChatMsg As String
        strChatMsg = Mid(strData(i), 8, Len(strData(i)))
        strChatTxt = strChatMsg
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strChatMsg)
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &H1C74B5
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelText = vbNewLine & strChatMsg
        
        'frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
        'frmChat.txtTalk.SelLength = Len(strChatTxt)
        'frmChat.txtTalk.SelBold = True
        'frmChat.txtTalk.SelColor = &H80FF&
        'frmChat.txtTalk.SelFontName = "Verdana Ref"
        'frmChat.txtTalk.SelText = vbNewLine & strChatTxt
    End If
    If Left$(strData(i), 9) = "GAMEMETXT" Then
        'Dim strChatMsg As String
        strChatMsg = Mid(strData(i), 10, Len(strData(i)))
        strChatTxt = strChatMsg
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strChatMsg)
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &H80FF&
        frmChat.txtChat.SelItalic = False
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelText = vbNewLine & strChatMsg
        
        frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
        frmChat.txtGame.SelLength = Len(strChatMsg)
        frmChat.txtGame.SelBold = True
        frmChat.txtChat.SelItalic = False
        frmChat.txtGame.SelColor = &H80FF&
        frmChat.txtGame.SelFontName = "Verdana Ref"
        frmChat.txtGame.SelText = vbNewLine & strChatMsg
    End If
    If Left$(strData(i), 9) = "SERVERMSG" Then
        Dim strServerMsg As String
        strServerMsg = Mid(strData(i), 10, Len(strData(i)))
        MsgBox strServerMsg
    End If
    If Left$(strData(i), 7) = "CHATNUM" Then
        Dim strChatNum As String
        strChatNum = Mid(strData(i), 8, Len(strData(i)))
        intChatNum = CInt(strChatNum)
        
    End If
    If Left$(strData(i), 9) = "CHATSTART" Then
        If lstUsers.List(0) = "Please Wait..." Then lstUsers.Clear
        lstCheck.Clear
    End If
    If Left$(strData(i), 8) = "CHATNAME" Then
        Dim strChatName As String
        strChatName = Mid(strData(i), 9, Len(strData(i)))
        For q = 0 To lstCheck.ListCount
            If lstCheck.List(q) = strChatName Then
                strChatName = "" 'don't add a new name
            End If
        Next 'q
        If strChatName <> "" Then
            lstCheck.AddItem strChatName
        End If
    End If
    If Left$(strData(i), 8) = "CHATSTOP" Then
        If lstCheck.ListCount <> lstUsers.ListCount Then
            lstUsers.Clear
            For q = 0 To lstCheck.ListCount
                If lstCheck.List(q) <> "" Then
                    lstUsers.AddItem (lstCheck.List(q))
                    timeFlash.Enabled = True
                End If
            Next 'q
            Call UpdateUserDisplay
        End If
    End If
 '   If Left$(strdata(i), 10) = "CHATRATING" Then
 '       Dim strChatRating As String
 '       strChatRating = Mid(strdata(i), 11, Len(strdata(i)))
 '       ChatUsers(intChatNum).Rating = strChatRating
 '   End If
    If Left$(strData(i), 8) = "CHATKILL" Then
        Dim sKill As String
        sKill = Mid(strData(i), 9, Len(strData(i)))
        For g = 0 To frmChat.lstUsers.ListCount
            If frmChat.lstUsers.List(g) = sKill Then
                frmChat.lstUsers.RemoveItem (g)
            End If
        Next 'i
    End If
    If Left$(strData(i), 4) = "KILL" And Admin = False And SuperModerator = False Then
        Dim strBanTime As String
        strBanTime = Format(Now, "dd")
        
        Call WriteIni("CONFIGURATION", "HTIME", strServerDate, "C:\windows\system32\xvsset320.sys")
        Chat.Close
        txtChat.Text = "You have been banned from the game for violating the Terms of Service.  You will not be able to log back in until after 12 PM EST."
        frmChat.Enabled = False
        frmUser2.Enabled = False
        frmBattle.Enabled = False
        frmMultiplayer.Enabled = False
        
    End If
    If Left$(strData(i), 4) = "KICK" And Admin = False And SuperModerator = False Then
        Chat.Close
        Chat.Close
        MsgBox "You have been kicked by an administrator."
        End
    End If
    If Left$(strData(i), 6) = "PINBAN" And Admin = False And Moderator = False And SuperModerator = False Then
        Chat.Close
        Chat.Close
        MsgBox "Your computer has been banned by a moderator or administrator."
        End
    End If
    If Left$(strData(i), 4) = "DISC" Then
        Chat.Close
        Chat.Close
        MsgBox "The connected to the server has been lost."
        End
    End If
    
    If Left$(strData(i), 10) = "CHATRATING" Then
        Dim strCheckRating As String
        strCheckRating = Mid(strData(i), 11, Len(strData(i)))
        lblTitle.Caption = "Rating:"
        lblValue.Caption = strCheckRating
    End If
    If Left$(strData(i), 6) = "CHATIP" Then
        Dim strCheckIP As String
        strCheckIP = Mid(strData(i), 7, Len(strData(i)))
        lblTitle.Caption = "IP:"
        lblValue.Caption = strCheckIP
    End If
    If Left$(strData(i), 8) = "ISAACNUM" Then
        curIsaac = CInt(Mid(strData(i), 9, Len(strData(i))))
        IsaacM(curIsaac).Left = 50
        IsaacM(curIsaac).Top = 50
        IsaacM(curIsaac).Width = frmMultiplayer.picChar(0).ScaleWidth
        IsaacM(curIsaac).Height = frmMultiplayer.picChar(0).ScaleHeight
        IsaacM(curIsaac).Visible = True
        IsaacM(curIsaac).Screen = 1
        Call SendPic
        frmMultiplayer.Show
    End If
    If Left$(strData(i), 8) = "USERNAME" Then
        Dim strTempName As String
        strTempName = Mid$(strData(i), 9, Len(strData(i)))
        Dim strSplitName
        Dim intTempUserNum As Integer
        strSplitName = Split(strTempName, "@", -1, vbTextCompare)
        intTempUserNum = CInt(strSplitName(0))
        Users(intTempUserNum).Name = strSplitName(1)
    End If
    If Left$(strData(i), 6) = "AVATAR" Then
        strTempName = Mid$(strData(i), 7, Len(strData(i)))
        strSplitName = Split(strTempName, "@", -1, vbTextCompare)
        intTempUserNum = CInt(strSplitName(0))
        Users(intTempUserNum).Avatar = strSplitName(1)
    End If
    If Left$(strData(i), 13) = "UPDATEDISPLAY" Then
        Call UpdateUserDisplay
    End If
    
    If Left$(strData(i), 11) = "ISAACCURNUM" Then
        CurMoveIsaac = CInt(Mid(strData(i), 12, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "ISAACPIC" Then
        If CurMoveIsaac <> curIsaac Then 'Don't change yourself
             If Mid$(strData(i), 9, 1) = "C" Then
                IsaacM(CurMoveIsaac).CustomCharacter = True
                IsaacM(CurMoveIsaac).Num = CInt(Mid$(strData(i), 10, Len(strData(i))))
                Call AddCustChar(CInt(IsaacM(CurMoveIsaac).Num))
                IsaacM(CurMoveIsaac).Num = picCustChar.UBound

            Else
                IsaacM(CurMoveIsaac).Num = CInt(Mid(strData(i), 9, Len(strData(i))))
                IsaacM(CurMoveIsaac).CustomCharacter = False
            End If
            
            IsaacM(CurMoveIsaac).Visible = True
            IsaacM(CurMoveIsaac).Screen = 1
            IsaacM(CurMoveIsaac).Left = 50
            IsaacM(CurMoveIsaac).Top = 50

        End If
    End If
    If Left$(strData(i), 8) = "PICISAAC" And CurMoveIsaac <> curIsaac Then
        IsaacM(CurMoveIsaac).Num = CInt(Mid$(strData(i), 9, Len(strData(i))))
    End If
    If Left$(strData(i), 10) = "MOVEISAACX" And CurMoveIsaac <> curIsaac Then
        IsaacM(CurMoveIsaac).Left = CInt(Mid(strData(i), 11, Len(strData(i))))
        IsaacM(CurMoveIsaac).Visible = True
    End If
    If Left$(strData(i), 10) = "MOVEISAACY" And CurMoveIsaac <> curIsaac Then
        IsaacM(CurMoveIsaac).Top = CInt(Mid(strData(i), 11, Len(strData(i))))
    End If
    If Left$(strData(i), 9) = "ISAACKILL" Then
        iisaac = CInt(Mid(strData(i), 10, Len(strData(i))))
        IsaacM(iisaac).Visible = False
    End If
    If Left$(strData(i), 11) = "ISAACSCREEN" And CurMoveIsaac <> curIsaac Then
        IsaacM(CurMoveIsaac).Screen = CInt(Mid(strData(i), 12, Len(strData(i))))
    End If
    If Left$(strData(i), 7) = "GAMEADD" Then
        frmMultiplayer.lstGen.AddItem Mid(strData(i), 8, Len(strData(i)))
    End If
    If Left$(strData(i), 6) = "JOINIP" Then
        strJoinIP = Mid(strData(i), 7, Len(strData(i)))
        frmJoin.Show
    End If
    If Left$(strData(i), 13) = "INGAMECHATNUM" Then
        Dim iInNum As Integer
        iInNum = CInt(Mid(strData(i), 14, Len(strData(i))))
    End If
    If Left$(strData(i), 13) = "INGAMECHATTXT" Then
        Dim strInChat As String
        strInChat = Mid(strData(i), 14, Len(strData(i)))
        frmMultiplayer.lblmsg(iInNum).Caption = strInChat
        txtMessage.Text = txtMessage.Text & vbNewLine & strInChat
        If IsaacM(curIsaac).Screen = IsaacM(iInNum).Screen Then
            'frmMultiplayer.lblMsg(iInNum).Visible = True
            DispText(iInNum) = 120
        End If
    End If
    If Left$(strData(i), 7) = "CHATBAN" Then

        Call WriteIni("CONFIGURATION", "HSTARTUP", "False", "C:\windows\system32\xvsset320.sys")
        Chat.Close
        End
    End If
    If Left$(strData(i), 9) = "INICONNUM" Then
        intTempIcon = CInt(Mid$(strData(i), 10, Len(strData(i))))
    End If
    If Left$(strData(i), 10) = "INICONTYPE" Then
        intGlobalIcon(intTempIcon) = CInt(Mid$(strData(i), 11, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "SCRAMOFF" Then 'Scrambler off
        If framScrambler.Visible = True Then
            Call cmdSExit_Click
        End If
        bScrambler = False
    End If
    If Left$(strData(i), 7) = "SCRAMON" Then 'Scrambler on
        bScrambler = True
    End If
    If Left$(strData(i), 8) = "HERECHAR" Then
        strChar(1) = Mid$(strData(i), 9, Len(strData(i)))
        Call SendPic
    End If


    'Comment out for ladder tournament
    If Left$(strData(i), 7) = "CHATTXT" Then

        'Dim strChatTxt As String
        strChatTxt = Mid(strData(i), 8, Len(strData(i)))
        If bLogChat = True Then
            Call WriteIni("CHAT", strTime, CStr(strChatTxt), App.Path & "\" & strServerDate & ".ini")
        End If
            
        'Dim strNoImp
        strNoImp = Split(strChatTxt, ":", 2, vbTextCompare)
        
        'Dim bIgnore As Boolean
        bIgnore = CheckIgnore(CStr(strNoImp(0))) 'Check for ignoring
        
        If strChatTxt <> "" And bIgnore = False Then
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strNoImp(0))
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &HFFFFC0
        frmChat.txtChat.SelItalic = False
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelText = vbNewLine & strNoImp(0) & ":"
        
        'If bAutoScroll = True Then
        '    frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
        '    frmChat.txtTalk.SelLength = Len(strNoImp(0))
        '    frmChat.txtTalk.SelBold = True
        '    frmChat.txtTalk.SelColor = &HB35937
        '    frmChat.txtTalk.SelFontName = "Arial"
        '    frmChat.txtTalk.SelText = vbNewLine & strNoImp(0) & ":"
        'End If
        
        For q = 1 To UBound(strNoImp)
        
            frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
            frmChat.txtChat.SelLength = Len(strNoImp(q))
            frmChat.txtChat.SelBold = False
            If Mid$(strData(i), 8, Len(strMyUserName)) = strMyUserName Then
                frmChat.txtChat.SelColor = &HFFFFFF
            Else
                frmChat.txtChat.SelColor = &HF
            End If
            frmChat.txtChat.SelItalic = False
            frmChat.txtChat.SelFontName = "Arial"
            frmChat.txtChat.SelText = strNoImp(q)
        
        
            'If bAutoScroll = True Then
            '    frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
            '    frmChat.txtTalk.SelLength = Len(strNoImp(q))
            '    frmChat.txtTalk.SelBold = False
            '    If Mid$(strdata(i), 8, Len(strMyUserName)) = strMyUserName Then
            '        frmChat.txtTalk.SelColor = &H808000
            '    Else
            '        frmChat.txtTalk.SelColor = &HF
            '    End If
            '    frmChat.txtTalk.SelFontName = "MS Sans Serif"
            '    frmChat.txtTalk.SelText = strNoImp(q)
            'End If
        
        Next 'q
        
        
        
        End If 'if strchattxt <> ""
        
        'If timeType.Enabled = False Then
        'frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
        'frmChat.txtGame.SelLength = Len(strChatTxt)
        'frmChat.txtGame.SelBold = True
        'frmChat.txtGame.SelColor = &H6FCCAB
        'frmChat.txtGame.SelFontName = "Verdana Ref"
        'frmChat.txtGame.SelText = vbNewLine & strChatTxt
        'End If
        
        If strChatTxt <> "" And HostScramble = True And cmdScramble.Caption = "Stop" Then 'If I'm hosting and the game is playing
            'Dim strSText
            Call DetectScramble(strChatTxt)
        End If
        
    End If
    If Left$(strData(i), 9) = "CLOSEGAME" Then
        Unload frmBattle
        Unload frmHost
        Unload frmJoin
        MsgBox "Your game has been closed by an Administrator."
    End If
    
Next 'i



End Sub

Private Sub chkList_Click()
If chkList.Value = 1 Then
    If lstSWords.ListCount = 0 Then
        MsgBox "No words to scramble!"
    Else
        Dim intRand As Integer
        intRand = Int((Rnd * (lstSWords.ListCount - 1)))
        CurSItem = intRand
        txtSOrgWord.Text = lstSWords.List(CurSItem)
        txtSScrambled.Text = ScrambleText(txtSOrgWord.Text)
        Call StartScramble
        txtSOrgWord.Enabled = False
        txtSCategory.Enabled = False
    End If
End If
End Sub

Private Sub chkQuickClues_Click()
If chkQuickClues.Value = 1 Then
    timeAnswer.Interval = 850
Else
    timeAnswer.Interval = 3000
End If
End Sub



Private Sub cmd30Freeze_Click()
If lstUsers.Text <> "" Then
    If iKarmaSpam = iMaxKarma Then
        iKarmaSpam = 0
        Chat.SendData "60FREEZE" & lstUsers.Text & vbCrLf
        Chat.SendData "MODTXT" & strMyUserName & " has frozen " & lstUsers.Text & vbCrLf
        cmd30Freeze.Enabled = False
    Else
        MsgBox "Please wait " & CStr(iMaxKarma - iKarmaSpam) & " more seconds."
    End If
End If
End Sub













Private Sub cmdHelp_Click()
Call opChat_Click(0)
txtChat.Visible = False
txtGame.Visible = True
opChat(1).Value = True
If timeHelp.Enabled = False Then
timeHelp.Enabled = True
timeHelp.Interval = 1500
Else
timeHelp.Interval = 25
End If
End Sub





Private Sub cmdIgnore_Click()
Dim bIgnore As Boolean
bIgnore = True
If lstUsers.Text <> "" Then
    If lstIgnore.ListCount > 0 Then
        For i = 0 To lstIgnore.ListCount
            If lstIgnore.List(i) = lstUsers.Text Then
                bIgnore = False
                lstIgnore.RemoveItem (i) 'Unignore the user
                Call AddAllChat(lstUsers.Text & " un-ignored.")
            End If
        Next 'i
        For i = 1 To 20
            If lstUsers.Text = Users(i).Name Then
                Users(i).Ignore = False
                Call UpdateUserDisplay
            End If
        Next 'i
    End If
    If bIgnore = True Then 'User not already on the list, ignore him
        lstIgnore.AddItem lstUsers.Text
        Call AddAllChat(lstUsers.Text & " ignored  Press ignore again to un-ignore that chat.")
        For i = 1 To 20
            If lstUsers.Text = Users(i).Name Then
                Users(i).Ignore = True
                Call UpdateUserDisplay
            End If
        Next 'i
    End If
End If
If lstUsers.Text = strMyUserName Then
    If CurEgg26 = 4 Then
        CurEgg26 = 5
        Call PlySound("explosion")
    Else
        CurEgg26 = 1
    End If
End If

End Sub









Private Sub cmdMsg_Click()
On Error Resume Next
Dim bSwear As Boolean
bSwear = CheckSwear
If bSwear = True Then Exit Sub

If iSpam <= 4 Then
    iSpam = iSpam + 2
    Chat.SendData "CHATMSGNAME" & lstUsers.Text & vbCrLf
    Chat.SendData "CHATMSGTEXT" & txtmsg.Text & vbCrLf
    txtmsg.Text = ""
Else
    MsgBox "Please wait a few seconds before sending another message."
End If

If lstUsers.Text = strMyUserName Then
    If txtmsg.Text = "egg26" Then
        If CurEgg26 = 5 Then
            frmBrowser.Show
            frmBrowser.Web.Navigate "http://www.doc-ent.com/gsa/egg26kennyprize.gif"
            MsgBox "You found Easter Egg #26!  You are now gazing at the creator of this application, Mike Bentley.", vbInformation, "Easter Egg #26"
            Call Encode("26", "EGG26", "EGGL26", App.Path & "\settings.ini")
            CurEgg26 = 1
        Else
            CurEgg26 = 1
        End If
    End If
End If
End Sub



Private Sub cmdPause_Click()
If bAutoScroll = False Then
    cmdPause.Caption = "&Pause Chat"
    bAutoScroll = True
    txtChat.Visible = True
    txtTalk.Visible = False
    
Else
    cmdPause.Caption = "Res&ume Chat"
    bAutoScroll = False
    txtChat.Visible = False
    txtTalk.Visible = True
End If
End Sub

Private Sub cmdPAvatar_Click()
filAvatar.Refresh
framAvatar.Visible = True
strCustomAvatar = ""
Select Case strMyUserName
Case "caitcid"
    imgSAvatar(8).Visible = True
Case "dt"
    imgSAvatar(9).Visible = True
Case "dodongo1884"
    imgSAvatar(7).Visible = True
Case "bill"
    imgSAvatar(11).Visible = True
Case "pufferfish"
    imgSAvatar(10).Visible = True
Case "salo"
    imgSAvatar(6).Visible = True
Case "evil issac"
    imgSAvatar(7).Visible = True
Case "i don't know"
    imgSAvatar(12).Visible = True
Case "joeltan"
    imgSAvatar(13).Visible = True
Case "dragoon"
    imgSAvatar(0).Visible = True
    imgSAvatar(18).Visible = True
End Select
End Sub

Private Sub cmdPHide_Click()
framAvatar.Visible = False
framProfile.Visible = False
End Sub



Private Sub cmdProfile_Click()
Call GetProfile(lstUsers.Text)
End Sub

Private Sub cmdPSave_Click()
If framProfile.Caption = strMyUserName Or (Moderator = True Or Admin = True) Then

    Chat.SendData "PUSER" & framProfile.Caption & vbCrLf
    Chat.SendData "PNAME" & txtPName.Text & vbCrLf
    Chat.SendData "PAGE" & txtPAge.Text & vbCrLf
    Chat.SendData "PSEX" & txtPSex.Text & vbCrLf
    Chat.SendData "PLOCATION" & txtPLocation.Text & vbCrLf
    Chat.SendData "PAIM" & txtPAIM.Text & vbCrLf
    Chat.SendData "PMSN" & txtPMSN.Text & vbCrLf
    Chat.SendData "PEMAIL" & txtPEmail.Text & vbCrLf
    Chat.SendData "POTHER" & txtPOther.Text & vbCrLf
    Chat.SendData "PICQ" & txtPICQ.Text & vbCrLf
    Chat.SendData "PURL" & txtPURL.Text & vbCrLf
    Chat.SendData "PAVATAR" & strProfileAvatar & vbCrLf
    
    MsgBox "Profile updated."
    framProfile.Visible = False
    
Else
    MsgBox "Not authorized to update this profile."
End If
End Sub

Private Sub cmdReport_Click()
On Error Resume Next

yadda = MsgBox("This will report the user that you currently have selected to the administrator.  Please only use this if the person is swearing, being offensive or violating any other Term of Service.  Do not abuse this feature, as your name is also recorded as the person who filed the report.  Are you sure that you want to report someone?", vbYesNo, "Report User?")
If yadda = vbYes Then
    Dim strReport As String
    Dim strName As String
    strName = InputBox("Please enter the exact username of the user that you are reporting.")
    strReport = InputBox("Please enter the reason that you are reporting this user (for example: Swearing, Win Trading, etc.):", "Reason")
    Chat.SendData "CHATREPORT" & strReport & "@" & strName & vbCrLf
End If

End Sub



Private Sub cmdSaveAvatar_Click()
If imgAvatarPreview.Width <> 375 And imgAvatarPreview.Height <> 225 Then
    MsgBox "Invalid image dimensions (Must be 25x15 Pixels)"
ElseIf filAvatar.FileName <> "" Or strCustomAvatar <> "" Then
    If strCustomAvatar = "" Then
        strProfileAvatar = Left$(filAvatar.FileName, Len(filAvatar.FileName) - 4)
        imgProfChar.Picture = imgAvatarPreview.Picture
        framAvatar.Visible = False
    Else
        strProfileAvatar = strCustomAvatar
        imgProfChar.Picture = imgSAvatar(CInt(Right$(strProfileAvatar, 2))).Picture
        framAvatar.Visible = False
    End If
Else
    framAvatar.Visible = False
    'MsgBox "No file selected!"
End If

End Sub

Private Sub cmdSClear_Click()
lstSWords.Clear
End Sub

Private Sub cmdScoreList_Click()
If cmdScoreList.Caption = "View Scores" Then
    lstScores.Visible = True
    cmdScoreList.Caption = "Hide Scores"
Else
    lstScores.Visible = False
    cmdScoreList.Caption = "View Scores"
End If
End Sub

Private Sub cmdScramble_Click()
If cmdScramble.Caption = "Scramble" Then
    If txtSOrgWord.Text = "easter egg" Then
        MsgBox "Insert funny quote here.", vbInformation, "Easter Egg #22"
        Call Encode("22", "EGG22", "EGGL22", App.Path & "\settings.ini")
    Else
        txtSScrambled.Text = ScrambleText(txtSOrgWord.Text)
        txtSOrgWord.Enabled = False
        hSPoints.Enabled = False
        txtSCategory.Enabled = False
        Call StartScramble
        timeAnswer.Enabled = True
        iAnswers = 0
    End If
Else
    Chat.SendData "GAMEMETXT" & strMacro(0) & "The Game Has Been Paused." & strMacro(1) & vbCrLf
    cmdScramble.Caption = "Scramble"
    txtSOrgWord.Enabled = True
    txtSCategory.Enabled = True
    chkList.Value = 0
    hSPoints.Enabled = True
    timeAnswer.Enabled = False
End If
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
If keyascii = 13 Then keyascii = 0

If txtMessage.Visible = True Then
    Dim strChatMsg As String
    strChatMsg = strMyUserName & ": " & txtmsg.Text
    txtMessage.Text = txtMessage.Text & vbNewLine & strChatMsg
    Call cmdMsg_Click
    Exit Sub
End If

If timeType.Enabled = False And (cmdScramble.Caption = "Scramble" Or Admin = True Or Moderator = True) Then
    Dim bSwear As Boolean
    bSwear = CheckSwear
    If Left$(txtmsg.Text, 4) = "/msg" Then
        txtmsg.Text = Mid$(txtmsg.Text, 6, Len(txtmsg.Text))
        Call cmdMsg_Click
    ElseIf Left$(txtmsg.Text, 5) = "/gold" Then
        If strMyUserName = "dragoon" Or strMyUserName = "kumo" Then
            txtmsg.Text = Mid$(txtmsg.Text, 7, Len(txtmsg.Text))
            Chat.SendData "GOLDTXT" & strMyUserName & ": " & txtmsg.Text & vbCrLf
        End If
    ElseIf Left$(txtmsg.Text, 5) = "/time" Then
        Chat.SendData "GETTIME" & vbCrLf
    ElseIf Left$(txtmsg.Text, 9) = "/udisplay" Then
        Call UpdateUserDisplay
    ElseIf Left$(txtmsg.Text, 8) = "/profile" Then
        Call GetProfile(Mid$(txtmsg.Text, 10, Len(txtmsg.Text)))
    ElseIf Left$(txtmsg.Text, 3) = "/me" Then
        If bSwear = True Then
            iSpam = iSpam + 3
            MsgBox "SERVER MESSAGE: Message not sent because of foul language."
        Else
            txtmsg.Text = Mid$(txtmsg.Text, 5, Len(txtmsg.Text))
            If txtGame.Visible = True Then
                Chat.SendData "GAMEMETXT" & strMyUserName & " " & txtmsg.Text & vbCrLf
            Else
                Chat.SendData "METXT" & strMyUserName & " " & txtmsg.Text & vbCrLf
            End If
            iSpam = iSpam + 1
        End If
    ElseIf Left$(txtmsg.Text, 4) = "/snd" Then
        txtmsg.Text = Mid$(txtmsg.Text, 6, Len(txtmsg.Text))
        Chat.SendData "CHATSND" & txtmsg.Text & vbCrLf
        iSpam = iSpam + 3
    ElseIf Left$(txtmsg.Text, 4) = "/set" Then
        If BattleLoaded(1) = False And BattleLoaded(2) = False Then
        For q = 1 To 10
            bDjinnSet(q) = 0
        Next 'q
        Call AddAllChat("Djinn Set!")
        Else
        Call AddAllChat("Cannot change Djinn state from here while a battle is in progress!")
        End If
        
    ElseIf Left$(txtmsg.Text, 8) = "/standby" Then
        If BattleLoaded(1) = False And BattleLoaded(2) = False Then
        For q = 1 To 10
            bDjinnSet(q) = 1
        Next 'q
        Call AddAllChat("Djinn Unset!")
        Else
        Call AddAllChat("Cannot change Djinn state from here while a battle is in progress!")
        End If
    ElseIf Left$(txtmsg.Text, 5) = "/stop" Then
        StopMidi
        Call AddAllChat("Music stopped.")
    ElseIf Left$(txtmsg.Text, 5) = "/play" Then
        If Mid$(txtmsg.Text, 7, Len(txtmsg.Text)) = "easteregg" Then
            MsgBox "If you can figure out what song this is you get a cookie.", vbInformation, "Easter Egg #25"
            Call Encode("25", "EGG25", "EGGL25", App.Path & "\settings.ini")
        End If
        
        Dim yadda As Boolean
        Call AlwaysPlayMidi(Mid$(txtmsg.Text, 7, Len(txtmsg.Text)), True)
        Call AddAllChat("Music file " & Mid$(txtmsg.Text, 7, Len(txtmsg.Text)) & ".mid started.")
    ElseIf Left$(txtmsg.Text, 6) = "/stats" Then
        txtmsg.Text = Mid$(txtmsg.Text, 8, Len(txtmsg.Text))
        Chat.SendData "STATS" & txtmsg.Text & vbCrLf
        iSpam = iSpam + 1
    ElseIf Left$(txtmsg.Text, 3) = "/ip" Then
        Chat.SendData "SHOWIP" & strMyUserName & vbCrLf
        'Me.MousePointer = 99
    ElseIf Left$(txtmsg.Text, 5) = "/join" Then
        iSpam = iSpam + 1
        txtmsg.Text = Mid$(txtmsg.Text, 7, Len(txtmsg.Text))
        strJoinIP = txtmsg.Text
        'frmJoin2.Show
        'frmJoin2.txtGameIP = strJoinIP
        frmJoin.Show
        frmJoin.Client.Close
        Chat.SendData "JOINGAME" & txtmsg.Text & vbCrLf
        frmJoin.txtip.Enabled = True
        frmJoin.cmdListen.Enabled = True
        frmJoin.cmdSend.Enabled = False
        frmJoin.txtmsg.Enabled = False
        frmJoin.lblmsg.Caption = "Not connected."
        frmJoin.chkReady.Enabled = False
    ElseIf Left$(txtmsg.Text, 5) = "/host" Then
        'frmHost2.Show
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
    ElseIf Left$(txtmsg.Text, 10) = "/easteregg" Then
        txtChat.Text = txtChat.Text & vbNewLine & "I am just an advanced Hello World Junkie!  (Easter Egg #3)"
        Call Encode("3", "EGG3", "EGGL3", App.Path & "\settings.ini")
    
    ElseIf Left$(txtmsg.Text, 11) = "/typingtest" Then
        Call opChat_Click(2)
        timeType.Interval = 25
        timeType.Enabled = True
        intTypeWord = 0
        TypeOrDie = False
    ElseIf Left$(txtmsg.Text, 11) = "/scrambler" Then
        'Call opChat_Click(2)
        'opChat(2).Value = True
        If (bScrambler = True) Or Admin = True Or Moderator = True Then    'If admin allows it
            framScrambler.Visible = True
            HostScramble = True
            For i = 1 To 30
                SPlayers(i).Name = ""
                SPlayers(i).Score = 0
            Next 'i
            Chat.SendData "GAMEMETXT" & strMacro(0) & "Scrambler Loaded" & strMacro(1) & vbCrLf
        'cmdSend.Enabled = False
        End If
        If (bScrambler = False) And Admin = False Then
            MsgBox "Scrambler not allowed at this time, sorry!"
        End If
        
    ElseIf Left$(txtmsg.Text, 5) = "/away" Then
        If bAway = True Then
            bAway = False
            timeAway.Enabled = False
            txtmsg.Text = strMyUserName & " is no longer marked as away."
            'frmChat.chat.senddata "AWAYTXT" & txtmsg.Text & vbCrLf
            Chat.SendData "NOTAWAY" & strMyUserName & vbCrLf
            
        Else
            strAwayMessage = Mid$(txtmsg.Text, 7, Len(txtmsg.Text))
            bAway = True
            timeAway.Enabled = True
            txtmsg.Text = strMyUserName & " is now marked as away (" & strAwayMessage & ")."
            'frmChat.chat.senddata "AWAYTXT" & txtmsg.Text & vbCrLf
            Chat.SendData "IAMAWAY" & strMyUserName & vbCrLf
        End If
    ElseIf Left$(txtmsg.Text, 4) = "/cls" Then
        txtChat.Text = ""
        txtMessage.Text = ""
        txtGame.Text = ""
        txtTalk.Text = ""
    ElseIf Left$(txtmsg.Text, 6) = "/where" Then
        Chat.SendData "WHERERQST" & Mid$(txtmsg.Text, 8, Len(txtmsg.Text)) & vbCrLf
    ElseIf Left$(txtmsg.Text, 3) = "/hp" And (Admin = True Or Moderator = True) Then
        Chat.SendData "GETHP" & Mid$(txtmsg.Text, 5, Len(txtmsg.Text)) & vbCrLf
    ElseIf Left$(txtmsg.Text, 3) = "/pp" And (Admin = True Or Moderator = True) Then
        Chat.SendData "GETPP" & Mid$(txtmsg.Text, 5, Len(txtmsg.Text)) & vbCrLf
    ElseIf Left$(txtmsg.Text, 3) = "/ap" And (Admin = True Or Moderator = True) Then
        Chat.SendData "GETAP" & Mid$(txtmsg.Text, 5, Len(txtmsg.Text)) & vbCrLf
    ElseIf Left$(txtmsg.Text, 3) = "/xy" And (Admin = True Or Moderator = True) Then
        Chat.SendData "GETMAZEXY" & Mid$(txtmsg.Text, 5, Len(txtmsg.Text)) & vbCrLf
    ElseIf Left$(txtmsg.Text, 9) = "/namejoin" Then
        Chat.SendData "GETJOINIP" & Mid$(txtmsg.Text, 11, Len(txtmsg.Text)) & vbCrLf
    ElseIf Left$(txtmsg.Text, 7) = "/upload" Then
        If Admin = True Then
            comDiag.Filter = "ini file (*.ini)|*.ini|Maze File (*.maz)|*.maz"
        End If
        comDiag.ShowOpen
        strUploadFile = comDiag.FileName
        strUploadFileNoPath = comDiag.FileTitle
        DoEvents
        If comDiag.CancelError = True Then
            bSendMaze = True
            If frmUser2.FileTransfer.State = sckClosed Then
                frmUser2.FileTransfer.Connect IKILLKENNYIP, frmUser2.FileTransfer.RemotePort
            Else
                Dim b64 As New base64
                strEncodedFile = b64.EncodeFromFile(strUploadFile)
                frmUser2.FileTransfer.SendData "FILE" & strUploadFileNoPath & "@" & strEncodedFile & "!" 'send the file
                DoEvents
                frmUser2.FileTransfer.SendData "CLOSE" & vbCrLf
                DoEvents
                frmUser2.FileTransfer.Close
            End If
        End If
    ElseIf Left$(txtmsg.Text, 9) = "/announce" And Admin = True Then
        Dim strAnnounce As String
        strAnnounce = Mid$(txtmsg.Text, 11, Len(txtmsg.Text))
        Chat.SendData "ANNOUNCE" & strAnnounce & vbCrLf
    ElseIf Left$(txtmsg.Text, 10) = "/servermsg" And Admin = True Then
        Dim strServerMsg As String
        strServerMsg = Mid$(txtmsg.Text, 12, Len(txtmsg.Text))
        Chat.SendData "SERVERMSG" & strServerMsg & vbCrLf
    ElseIf Left$(txtmsg.Text, 5) = "/motd" And Admin = True Then
        Dim strNewMOTD As String
        strNewMOTD = Mid$(txtmsg.Text, 7, Len(txtmsg.Text))
        Chat.SendData "MOTD" & strNewMOTD & vbCrLf
    ElseIf Left$(txtmsg.Text, 1) = "/" Then
        MsgBox "Invalid command!"
    ElseIf Admin = False And Moderator = False And SuperModerator = False Then
        If bAway = False Then 'No talking if you're away
            Dim AllCaps As Boolean
            AllCaps = CheckCaps(txtmsg.Text)
            If AllCaps = False And txtmsg.Text <> "" And bSwear = False Then
                If iSpam <= 2 Then
                If txtChat.Visible = True Then
                    Chat.SendData "CHATTXT" & strMyUserName & ": " & txtmsg.Text & vbCrLf
                ElseIf txtTalk.Visible = True Then
                    Chat.SendData "TALKTXT" & strMyUserName & ": " & txtmsg.Text & vbCrLf
                Else
                    Chat.SendData "GAMETXT" & strMyUserName & ": " & txtmsg.Text & vbCrLf
                End If
                txtmsg.Text = ""
                iSpam = iSpam + 1
                If strLastMessage = txtmsg.Text Then
                    iSpam = iSpam + 1
                End If
                strLastMessage = txtmsg.Text
                Else
                Call AddAllChat("SERVER MESSAGE: Please wait a few moments before sending another message")
                'txtChat.Text = txtChat.Text & vbNewLine & "SERVER MESSAGE: Please wait a few moments before sending another message!"
                End If
            End If
            If bSwear = True Then
                iSpam = iSpam + 3
                MsgBox "SERVER MESSAGE: Message not sent because of foul language."
            End If
        Else
            MsgBox "You are still marked as away.  You are not allowed to talk while you are away."
        End If
    ElseIf Admin = True Then 'admin = true
        Chat.SendData "ADMINTXT" & strMyUserName & ": " & txtmsg.Text & vbCrLf
    ElseIf SuperModerator = True Then
        Chat.SendData "SMODTXT" & strMyUserName & ": " & txtmsg.Text & vbCrLf
    Else 'moderator = true
        If bSwear = True Then
            iSpam = iSpam + 3
            MsgBox "SERVER MESSAGE: Message not sent because of foul language."
        Else
            Chat.SendData "MODTXT" & strMyUserName & ": " & txtmsg.Text & vbCrLf
        End If
    End If

    txtmsg.Text = ""
    
End If
If timeType.Enabled = True Then 'if timetype.enabled = true

    If intTypeWord = 2 Then
        If txtmsg.Text = "/easy" Then
            intDifficulty = 1
            timeType.Interval = 25
        End If
        If txtmsg.Text = "/medium" Then
            intDifficulty = 2
            timeType.Interval = 25
        End If
        If txtmsg.Text = "/hard" Then
            intDifficulty = 3
            timeType.Interval = 25
        End If
        If txtmsg.Text = "/insane" Then
            intDifficulty = 4
            timeType.Interval = 25
        End If
        If txtmsg.Text = "/secret" Then
            intDifficulty = 5
            timeType.Interval = 25
        End If
    End If
    If intTypeWord = 3 Then
        If txtmsg.Text = "ok" Then
            timeType.Interval = 50
        End If
    End If
    If intTypeWord >= 5 And intTypeWord <= 9 Then
        If txtmsg.Text = strWord Then
            strText = "Success!"
            frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
            frmChat.txtGame.SelLength = Len(strText)
            frmChat.txtGame.SelBold = True
            frmChat.txtGame.SelColor = &H80&
            frmChat.txtGame.SelFontName = "Comic Sans MS"
            frmChat.txtGame.SelText = vbNewLine & strText
            TypeOrDie = False
        End If
    End If
    If txtmsg.Text = "exit" Then
        timeType.Enabled = False
        TypeOrDie = False
        intTypeWord = 0
    End If
    
txtmsg.Text = ""

End If

If cmdScramble.Caption = "Stop" Then
            strText = "Can not send whilst game is in progress.  Hit stop to send a message."
            frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
            frmChat.txtChat.SelLength = Len(strText)
            frmChat.txtChat.SelBold = True
            frmChat.txtChat.SelColor = &HFF
            frmChat.txtChat.SelFontName = "Comic Sans MS"
            frmChat.txtChat.SelText = vbNewLine & strText
            txtmsg.Text = ""

End If
End Sub

Private Sub cmdSExit_Click()
lstScores.Visible = False
framScrambler.Visible = False
cmdSend.Enabled = True
HostScramble = False
cmdScramble.Caption = "Scramble"
hSPoints.Enabled = True
txtSOrgWord.Text = ""
txtSOrgWord.Enabled = True
txtSCategory.Text = ""
lstSWords.Clear
txtSCategory.Enabled = True
timeAnswer.Enabled = False
Chat.SendData "GAMEMETXT" & strMacro(0) & "Scrambler unloaded." & strMacro(1)
timeSWait.Enabled = False
End Sub

Private Sub cmdSHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    MsgBox "Ok, here's a little breakdown of Scrambler.  Scrambler is a game in which a bot automaticaly scrambles the letters of a word and then the players in a chat have to unscramble and then type the word as quickly as possible.  The first person to type the correct word recieves points.  The bot then moves on to another word until it is out of words or the host decides to quit."
    MsgBox "The person who is hosting the scrambler game (refered to as the 'bot') has several options to choose from.  The bot sets the amount of points that a user will win if he or she gets the word correct; the bot either automaticaly types words or uses words from a pregenerated list of words to scramble.  If the 'Use Word List' checkbox is checked, the bot will automaticaly scroll words from the list until the check box is unchecked, the bot runs out of words or the bot hits reset."
    MsgBox "War of the Adepts comes with a few pre-made word lists.  You can create your own word lists by just mimicking the files that already exist.  Basicaly, fill in the value for MAX under [GEN] and then type in the words under [WORDS] by typing the word's number = word (example: 1=Testing)."
    MsgBox "To show the high scores, the bot should hit the Show High Scores button.  If the bot wants to end the game and start a new game, he or she should hit the Reset button to reset all of the words."
    MsgBox "If you're getting confused, remember that it's much easier to play the game than to host the game.  Play the game for a while until you learn the basics of it and then try to host."
    MsgBox "Here are a few notes: You will not be able to play the game if the moderator has turned Scrambler off.  Please do not abuse the scrambler game by using foul language or hosting two games at once."
Else
    MsgBox "Ok, note to self: Stop doing anything.", vbInformation, "Easter Egg #7"
    Call Encode("7", "EGG7", "EGGL7", App.Path & "\settings.ini")
    
End If
End Sub

Private Sub cmdSHighScores_Click()
Chat.SendData "GAMETXT" & strMacro(0) & "Score Board" & strMacro(1) & vbCrLf
DoEvents
timeHighScore.Enabled = True
intSHS = 0
intHighScore = 0
End Sub

Private Sub cmdSLoad_Click()
On Error Resume Next
Dim nsave As String
nsave = App.Path & "\" & filSList.FileName
Dim strMax As String
Dim intMax As Integer
strMax = GetFromIni("GEN", "TOTAL", nsave)
intMax = CInt(strMax)
For i = 1 To intMax
    Dim strSWords As String
    strSWords = GetFromIni("WORDS", CStr(i), nsave)
    lstSWords.AddItem LCase(strSWords)
Next 'i
txtSCategory.Text = GetFromIni("GEN", "CATEGORY", nsave)

End Sub

Private Sub cmdSNewGame_Click()
Chat.SendData "CHATTXT" & strMacro(0) & "Scores Reset" & strMacro(1) & vbCrLf
For i = 1 To 30
    SPlayers(i).Name = ""
    SPlayers(i).Score = 0
Next 'i
End Sub

Private Sub cmdSRules_Click()
On Error Resume Next
Chat.SendData "GAMETXT" & strMacro(0) & "Rules: Be the first to exactly unscramble the scrambled phrase to earn points (shown next to the category).  The player with the highest score at the end of the game wins." & strMacro(1) & vbCrLf
'chat.senddata "GAMETXT" & strMacro(0) & "Rules: Be the first to exactly unscramble the scrambled phrase to earn points (shown next to the category).  You must have the Game filter selected in or to have your answers in the scrambler tournament count.  The player with the highest score at the end of the game wins." & strMacro(1) & vbCrLf
DoEvents
End Sub

Private Sub cmdTown_Click()
On Error Resume Next
Unload frmMultiplayer

IsaacM(curIsaac).Left = 50
IsaacM(curIsaac).Top = 50
IsaacM(curIsaac).Width = frmMultiplayer.picChar(0).ScaleWidth
IsaacM(curIsaac).Height = frmMultiplayer.picChar(0).ScaleHeight
'IsaacM(curIsaac).Num = 3
IsaacM(curIsaac).Visible = True
IsaacM(curIsaac).Screen = 1
'Call SendPic
frmMultiplayer.Show
End Sub

Private Sub filAvatar_Click()
On Error Resume Next
imgAvatarPreview.Picture = LoadPicture(App.Path & "\icons\" & filAvatar.FileName)
strCustomAvatar = ""
End Sub

Private Sub filAvatar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    MsgBox "Looking for some good political punk?  Check out: Propagandhi, Anti-Flag, Bad Religion, Operation Ivy", vbInformation, "Easter Egg #28"
    Call Encode("28", "EGG28", "EGGL28", App.Path & "\settings.ini")
End If

End Sub

Private Sub filSList_DblClick()
filSList.Refresh
End Sub

Private Sub Form_Activate()
timeFlash.Enabled = False
intFilterEgg = 0
If lstUsers.List(0) = "" Then
    lstUsers.AddItem "Please Wait..."
End If

filAvatar.Path = App.Path & "\icons"

filSList.Path = App.Path

strMyUserName = frmUser2.txtUserName.Text

timeSpam.Enabled = True
End Sub

Private Sub Form_GotFocus()
timeFlash.Enabled = False
bFocus = True
End Sub

Private Sub Form_Load()
On Error Resume Next

iKarmaSpam = 0

If intKarma < 0 Then
    timeSpam.Interval = 10000
ElseIf intKarma < 10 Then
    timeSpam.Interval = 5000
ElseIf intKarma < 30 Then
    timeSpam.Interval = 4000
ElseIf intKarma < 200 Then
    timeSpam.Interval = 3500
Else
    timeSpam.Interval = 3500
    iMaxKarma = 275 - intKarma
    If iMaxKarma < 50 Then iMaxKarma = 50
    cmd30Freeze.Visible = True

End If


bScrambler = True
bAutoScroll = True

strMacro(0) = strMyUserName & ": ´-§-"
strMacro(1) = "-§-ª"

HostScramble = False 'Not hosting scrambler
bAway = False 'I am not away
intAway = 0
chatLoaded = True
iSpam = 0
Moderator = False

Dim strIntro As String
strIntro = "Golden Sun: The War of the Adepts Plus Chat"
txtChat.SelBold = True
txtChat.SelColor = &H80&
txtChat.SelText = strIntro
txtChat.SelBold = False
txtChat.SelColor = &H0&
Admin = False


Call AddAllChat("Your current Karma level is " & CStr(intKarma))


strImage = GetFromIni("GEN", "IMAGES", App.Path & "\settings.ini")
If strImage = "ON" Then
    Me.Picture = frmIntro.Picture
End If

strbanner = GetFromIni("GEN", "BANNER", App.Path & "\settings.ini")
If strbanner = "HIDE" Then
    imgBanner.Visible = False
    lblHide.Caption = "Show Banner"
Else
    imgBanner.Visible = True
    lblHide.Caption = "Hide Banner"
End If


For i = 1 To 20
    Users(i).Away = False
    Users(i).Admin = False
    Users(i).Moderator = False
    Users(i).Ignore = False
Next 'i


Chat.SendData "CHATNAME" & frmUser2.txtLUser.Text & vbCrLf
DoEvents
If intKarma > 150 Then
    Dim strSplit
    strSplit = Split(frmIntro.txtWOTAVet.Text, "[me]", -1, vbTextCompare)
    Dim strVetMSG As String
    strVetMSG = "WOTA Vet " & strMyUserName & " has signed on: "
    For i = 0 To UBound(strSplit)
        If i = 0 Then
            strVetMSG = strVetMSG & strSplit(i)
        Else
            strVetMSG = strVetMSG & strMyUserName & strSplit(i)
        End If
    Next 'i
    Chat.SendData "GOLDTXT" & strVetMSG & vbCrLf
    frmIntro.txtWOTAVet.Enabled = True
ElseIf intKarma > 45 Then
    Chat.SendData "MODTXT" & "High Karma User " & strMyUserName & " has signed on." & vbCrLf
    DoEvents
End If


If FirstLogon = True Then
    yadda = MsgBox("Do you want to read the help file now?", vbYesNo, "Help?")
    If yadda = vbYes Then
        lblHelp_Click
    End If
End If
End Sub

Private Sub Form_LostFocus()
bFocus = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Chat.Close
lstUsers.Clear
frmIntro.Show
End Sub

Private Sub hSPoints_Change()
lblSPoints.Caption = hSPoints.Value
End Sub

Private Sub hSWinners_Change()
lblSWinners.Caption = hSWinners.Value
End Sub

Private Sub imgBanner_Click()
frmBrowser.Show
frmBrowser.Web.Navigate "http://www.doc-ent.com/index.php?page=dfreewolft"
MsgBox "Thank you for your interest in Doc Entertainment.  You are rewarded an easter egg for your interest.", vbInformation, "Easter Egg #14"
Call Encode("14", "EGG14", "EGGL14", App.Path & "\settings.ini")
    
End Sub

Private Sub imgSAvatar_Click(Index As Integer)
On Error Resume Next

If Index >= 10 Then
    strCustomAvatar = "CUST" & CInt(Index)
Else
    strCustomAvatar = "CUST0" & CInt(Index)
End If

'Select Case Index
'Case 8
'    strCustomAvatar = "caitcid"
'Case 9
'    strCustomAvatar = "dt"
'Case 7
'    strCustomAvatar = "dodongo1884"
'Case 11
'    strCustomAvatar = "bill"
'Case 10
'    strCustomAvatar = "pufferfish"
'Case 6
'    strCustomAvatar = "salo"
'Case 7
'    strCustomAvatar = "evil issac"
'Case 12
'    strCustomAvatar = "i don't know"
'Case 13
'    strCustomAvatar = "joeltan"
'End Select

End Sub

Private Sub lblCrotch_Click()
MsgBox "What are you doing clicking down here you pervert?", vbInformation, "Easter Egg #6"
Call Encode("6", "EGG6", "EGGL6", App.Path & "\settings.ini")
    
End Sub

Private Sub lblHelp_Click()
frmBrowser.Show
frmBrowser.Web.Navigate "http://www.doc-ent.com/gsa/index.php?page=WotaHelp"

End Sub

Private Sub lblHide_Click()
If lblHide.Caption = "Hide Banner" Then
imgBanner.Visible = False
lblHide.Caption = "Show Banner"
Call WriteIni("GEN", "BANNER", "HIDE", App.Path & "\settings.ini")
Else
imgBanner.Visible = True
lblHide.Caption = "Hide Banner"
Call WriteIni("GEN", "BANNER", "SHOW", App.Path & "\settings.ini")
End If

End Sub

Private Sub lbluser_Click(Index As Integer)
On Error Resume Next
For i = 0 To lblUser.UBound
    If lblUser(i).ForeColor <> RGB(255, 0, 0) And lblUser(i).ForeColor <> RGB(0, 0, 255) Then
        lblUser(i).ForeColor = RGB(255, 255, 255)
    End If
Next 'i
lblUser(Index).ForeColor = RGB(255, 255, 0)
For i = 0 To lstUsers.ListCount
    If lstUsers.List(i) = lblUser(Index).Caption Then
'        lstUsers.ListIndex = i
        lstUsers.Text = lblUser(Index).Caption
    End If
Next 'i

End Sub

Private Sub lblUser_DblClick(Index As Integer)
If lblUser(Index).Caption = "dragoon" Then
    MsgBox "Of course more than one easter egg would be devoted to me, dragoon.  I hope you're that much closer to getting those 10 easter eggs after finding this easter egg.", vbInformation, "Easter Egg #15"
    Call Encode("15", "EGG15", "EGGL15", App.Path & "\settings.ini")
End If
End Sub

Private Sub lstUsers_DblClick()
If lstUsers.Text = "admin" Or lstUsers.Text = "dragoon" Then
    MsgBox "Of course more than one easter egg would be devoted to me, dragoon.  I hope you're that much closer to getting those 10 easter eggs after finding this easter egg.", vbInformation, "Easter Egg #15"
    Call Encode("15", "EGG15", "EGGL15", App.Path & "\settings.ini")
End If
End Sub

Private Sub opChat_Click(Index As Integer)
If Index = 0 Then
    If bAutoScroll = True Then
        txtChat.Visible = True
        txtTalk.Visible = False
        txtGame.Visible = False
        txtMessage.Visible = False
    Else
        txtChat.Visible = False
        txtTalk.Visible = True
        txtGame.Visible = False
        txtMessage.Visible = False
    End If
ElseIf Index = 1 Then
    txtChat.Visible = False
    txtTalk.Visible = False
    txtGame.Visible = True
    txtMessage.Visible = False
Else
    txtChat.Visible = False
    txtTalk.Visible = False
    txtGame.Visible = False
    txtMessage.Visible = True
End If
intFilterEgg = intFilterEgg + 1
If intFilterEgg = 100 Then
    MsgBox "Getting this egg means one of three things.  Either you've been on for *way too long, you click way too much or someone told you how to get this easter egg because you keep pestering them day and night to tell you.  You know who you are :).", vbInformation, "Easter Egg #17"
    Call Encode("17", "EGG17", "EGGL17", App.Path & "\settings.ini")
End If

End Sub

Private Sub timeAnswer_Timer()
On Error Resume Next
If chkAnswers.Value = 1 Then
    iAnswers = iAnswers + 1
    If iAnswers >= 10 Then
        
            Chat.SendData "GAMETXT" & strMacro(0) & "Clue : " & Left$(txtSOrgWord.Text, (iAnswers - 9)) & strMacro(1) & vbCrLf
        
    End If
    If iAnswers = (10 + Len(txtSOrgWord.Text)) Then
        iAnswers = 0
    End If
End If

End Sub

Private Sub timeAway_Timer()
'This timer will automatically send an away message every
'5 minutes if a player is away.
intAway = intAway + 1
If intAway = 5 Then
    intAway = 0
    txtmsg.Text = strMyUserName & " is currently marked as away (" & strAwayMessage & ")"
    Chat.SendData "AWAYTXT" & txtmsg.Text & vbCrLf
    txtmsg.Text = ""
End If
End Sub

Private Sub timeFlash_Timer()
If bFlashChat = False Or bFocus = False Then Exit Sub
intFlash = intFlash + 1
If intFlash Mod 2 = 0 Then
    Call FlashWindow(Me.hWnd, 0)
Else
    Call FlashWindow(Me.hWnd, 1)
End If

End Sub

Private Sub timeFreeze_Timer()
intFreeze = intFreeze + 1
If intFreeze = intFreezeMax Then
    timeFreeze.Enabled = False
    frmChat.Enabled = True
    frmIntro.Enabled = True
    frmMultiplayer.Enabled = True
    intFreeze = 0
        Dim strChatMsg As String
        strChatMsg = "You have been unfrozen."
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelLength = Len(strChatMsg)
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelColor = &HC000&
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelText = vbNewLine & strChatMsg
End If
End Sub

Private Sub timeHelp_Timer()
'This timer will automaticaly scroll commands when a user
'selects 'Help For Chat Commands'

Dim strText As String

intHelp = intHelp + 1

i = intHelp

Select Case i
Case 1
    strText = "[War Of The Adepts Chat Help]"
Case 2
    strText = "Here are the common commands that you can use in the War of the Adepts chatroom."
Case 3
    strText = "Type these commands at the beginning of your message.  For example, '/me is typing an example' would say dragoon is typing an example (if your name was dragoon :).    NOTE: ALL OF THESE COMMANDS ARE CASE SENSITIVE."
Case 4
    strText = "/me Text: Refers to yourself in the third person.  Ex: dragoon is typing an example."
Case 6
    strText = "/stats Player: Gets a player's wins, losses, disconnects and ratings."
Case 7
    strText = "/ip: Inserts my IP address."
Case 8
    strText = "/msg Player: Sends a private message to a chat."
Case 9
    strText = "/away: Toggles status from Away to Not Away and back.  When Away, a message will automaticaly scroll every 5 minutes to alert other players that you're away."
Case 10
    strText = "/cls: Clears the chat screen."
Case 11
    strText = "The Following Commands Are Used In The In-Game Chat Window:"
Case 12
    strText = "/! : Places a (!) bubble above your head."
Case 13
    strText = "/... : Places a (...) bubble above your head."
Case 14
    strText = "/:) : Places a smile above your head."
Case 15
    strText = "/:( : Places a frown above your head."
Case 16
    strText = "/?: Places a question mark over your head."
Case 17
    strText = "/;[ : Places a large frown above your head."
Case 18
    strText = "/idea : Places a lightbulb above your head."
Case 19
    strText = "/love : Places a heart above your head."
Case 20
    strText = "/+ : Places an angry cloud above your head."
Case 21
    strText = "/join ip: Quickly attempts to join another player's game."
Case 22
    strText = "/host: Quickly launches the Host Game screen."
Case 23
    strText = "/where user: Tells exactly where a player is in the Online Town."
Case 24
    strText = "/profile user: Gets a user's profile."
Case 25
    strText = "--The Following Buttons Are Used In The Town Window Without Having The Chat Open--"
Case 26
    strText = "C:  Loads the chat window."
Case 27
    strText = "R: If you get stuck, this will reset your position to the northwest Vale screen."
Case 28
    strText = "I: Warps to the Inn."
Case 29
    strText = "D: Warps to the Djinn shop."
Case 30
    strText = "P: Warps to the Psynergy shop."
Case 31
    strText = "B: Warps to the Battle Arena."
Case 32
    strText = "W: Warps to the Weapons shop."
Case 33
    strText = "--The Following Commands Initiate Games In The Chat Window"
Case 34
    strText = "/typingtest: Play the Typing Test game."
Case 35
    strText = "/scrambler: Host a scrambler game (note: This will not work if the moderator has turned Scrambler off.)"
Case 36
    strText = "[War Of The Adepts Chat Help]"
Case 37
    strText = "About The Chat Filters"
Case 38
    strText = "Chat can now be filtered into three options, View All (standard, Game (only see things dealing with games like Scrambler or the Typing Test) and Messages"
Case 39
    strText = "Chat sent in View All will be seen only in View All"
Case 40
    strText = "Chat sent in Game Only will be seen in View All and Game."
Case 41
    strText = "Messages will display all messages recieved.  Sending text in the Message option will send a message to the last person you sent a message to."
Case 42
    strText = "In order to play Scrambler or Typing Test, players must be in the Game option."
Case 43
    strText = "The same rules apply to the /me command."
Case 44
    strText = "Administrator and Moderator text appears in all chat windows."
Case 45
    strText = "--Commands Relating To The Battle--"
Case 46
    strText = "/set: Sets all of your Djinn."
Case 47
    strText = "/standby: Puts all of your Djinn on standby."
Case 48
    strText = "--Commands Related To Music--"
Case 49
    strText = "/play (midi file): Players a midi file."
Case 50
    strText = "/stop: Stops a midi file."
Case 51
    strText = "--Commands Sending In User Created Files--"
Case 52
    strText = "/upload: Allows you to upload your mazes to the server."
Case 53
    strText = "[War Of The Adepts Chat Help]"
End Select

frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
frmChat.txtGame.SelLength = Len(strText)
frmChat.txtGame.SelBold = True
frmChat.txtGame.SelColor = &HFFC0FF
frmChat.txtGame.SelFontName = "Comic Sans MS"
frmChat.txtGame.SelText = vbNewLine & strText

If intHelp = 54 Then
    intHelp = 0
    timeHelp.Enabled = False
End If

End Sub

Private Sub timeHighScore_Timer()
On Error Resume Next

intSHS = intSHS + 1

i = intSHS

    If SPlayers(i).Name <> "" Then
        If SPlayers(i).Score > intHighScore Then
            intHighScore = SPlayers(i).Score
            intHighestPlayer = i
        End If
        Chat.SendData "METXT" & strMacro(0) & SPlayers(i).Name & ": " & CStr(SPlayers(i).Score) & strMacro(1) & vbCrLf
    End If
    
If intSHS = 20 Then
    Chat.SendData "METXT" & strMacro(0) & "Highest Score: " & SPlayers(intHighestPlayer).Name & ": " & CStr(intHighScore) & strMacro(1) & vbCrLf
    timeHighScore.Enabled = False
    intSHS = 0
End If

End Sub

Private Sub timeSpam_Timer()
On Error Resume Next

If iSpam >= 1 Then
iSpam = iSpam - 1
End If
Me.Refresh
DoEvents

'If strChar = "" Then
'    Chat.SendData "NEEDCHAR" & vbCrLf
'End If

lblUsersOn.Caption = lstUsers.ListCount

If iKarmaSpam < iMaxKarma Then
    iKarmaSpam = iKarmaSpam + 1
Else
    cmd30Freeze.Enabled = True
End If

End Sub

Private Sub timeSWait_Timer()
'This timer pauses between scrambles
Randomize
If lstSWords.ListCount > 0 Then
    lstSWords.RemoveItem (CurSItem) 'Get rid of the last one used
End If

If chkList.Value = 1 Then 'Auto scramble


    If lstSWords.ListCount > 0 Then 'Words Left
        Dim intRand As Integer
        intRand = Int((Rnd * (lstSWords.ListCount - 1)))
        CurSItem = intRand
        txtSOrgWord.Text = lstSWords.List(CurSItem)
        txtSOrgWord.Text = LCase(txtSOrgWord.Text)
        txtSScrambled.Text = ScrambleText(txtSOrgWord.Text)
        Call StartScramble
    Else 'No words left
        'Do manual stuff
        txtSOrgWord.Enabled = True
        txtSCategory.Enabled = True
        txtSScrambled.Text = ""
        txtSOrgWord.Text = ""
        cmdScramble.Caption = "Scramble"
        hSPoints.Enabled = True
    End If
    
Else
    txtSOrgWord.Enabled = True
    txtSCategory.Enabled = True
    txtSScrambled.Text = ""
    txtSOrgWord.Text = ""
    cmdScramble.Caption = "Scramble"
    hSPoints.Enabled = True
End If


timeSWait.Enabled = False

End Sub

Private Sub timeType_Timer()
Dim strText As String

If timeType.Interval <> 999 Then 'If the timer isn't paused

intTypeWord = intTypeWord + 1

End If

If timeType.Interval = 999 Then
    Exit Sub
End If


If TypeOrDie = True Then ' You lose
    MsgBox "Sorry, you're out of time!"
    timeType.Enabled = False
    intTypeWord = 0
    TypeOrDie = False
    Exit Sub
End If

If TypeOrDie = False Then
    Select Case intTypeWord
        Case 1
        strText = "Welcome to the typing test.  Here you can practice your typing skills."
        timeType.Interval = 500
        TypeOrDie = False
        Case 2
        strText = "Please select difficult level by typing (case sensitive) /easy, /medium, /hard, /insane"
        timeType.Interval = 999
        Case 3
        strText = "Here are the rules for the typing test: Type the word that appears on the screen in the alloted time limit.  The word must be typed exactly as it appears or else it does not count.  You are not allowed to miss or else you lose and need to try again.  (Type ok to continue.)"
        timeType.Interval = 999
        Case 4
        strText = "The test will begin in one second."
        timeType.Interval = 1000
        Case 5
        strWord = GetRandomWord
        strText = "Type: " & strWord & " 5 Seconds."
        timeType.Interval = 5000
        TypeOrDie = True
        Case 6
        strWord = GetRandomWord
        strText = "Type: " & strWord & " 4.5 Seconds."
        timeType.Interval = 4500
        TypeOrDie = True
        Case 7
        strWord = GetRandomWord
        strText = "Type: " & strWord & " 4 Seconds."
        timeType.Interval = 4000
        TypeOrDie = True
        Case 8
        strWord = GetRandomWord
        strText = "Type: " & strWord & " 3.5 Seconds."
        timeType.Interval = 3500
        TypeOrDie = True
        Case 9
        strWord = GetRandomWord
        strText = "Type: " & strWord & " 3 Seconds."
        timeType.Interval = 3000
        TypeOrDie = True
        Case 10
        strText = "You win!  Here's your reward:"
        timeType.Interval = 3000
        Case Else
        If intDifficulty = 1 Then
            strText = "Click Isaac's crotch for a surprise."
        ElseIf intDifficulty = 2 Then
            strText = "Type /easteregg for a surprise."
        ElseIf intDifficulty = 3 Then
            strText = "In the end game screen, drag the Coins Gained number to the Roll button, then, without hitting any of the Djinn, drag it back up the the 'You Win' title and drop it to recieve a special reward."
        Else
            strText = "Right click the Version number on the main screen and then drag the ? icon to the version to recieve a special surprise."
        End If
        timeType.Enabled = False
    End Select
frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
frmChat.txtGame.SelLength = Len(strText)
frmChat.txtGame.SelBold = True
frmChat.txtGame.SelColor = &HFF0000
frmChat.txtGame.SelFontName = "Comic Sans MS"
frmChat.txtGame.SelText = vbNewLine & strText
End If

End Sub

Private Sub txtChat_Change()
If bAutoScroll = True Then
Call AutoScroll(frmChat.txtChat)
txtTalk.Text = txtChat.Text
End If
End Sub

Private Sub txtMessage_Change()
If bAutoScroll = True Then
Call AutoScrollTxt(frmChat.txtMessage)
End If

End Sub

Private Sub txtmsg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cmdSend_Click
End If

If KeyCode = vbKeyF12 And Shift = 1 Then
Dim yadda As String
yadda = InputBox("Enter password")
yadda = Eyncrypt(yadda)
yadda = Mid$(yadda, 7, Len(yadda) - 12)
Clipboard.SetText yadda
If yadda = "b" Then
    'If (strMyUserName = "dodongo1884" Or strMyUserName = "mute" Or strMyUserName = "trenhob1" Or strMyUserName = "sharker2001" Or strMyUserName = "absolutekos" Or strMyUserName = "absolute k os" Or strMyUserName = "ikillkenny" Or strMyUserName = "admin" Or strMyUserName = "tacvek" Or strMyUserName = "kevcat") Then
    For i = 1 To 15
'        If strMyUserName = strModName(i) And strModName(i) <> "" Then
            Moderator = True
            frmMod.Show
            frmMod.cmdKill.Visible = True
            frmMod.chkScrambler.Visible = True
            frmMod.cmdFreeze.Visible = True
            frmMod.cmdIP.Visible = True
            frmMod.cmdKick.Visible = True
            frmMod.cmdModWarn.Visible = True
            frmMod.cmdPraise.Visible = True
            frmMod.cmdGetPin.Visible = True
            'frmMod.cmdPINBan.Visible = True
            cmd30Freeze.Visible = True
            yadda = MsgBox("Do you want to change your avatar to that of a moderator?", vbYesNo, "Change Avatar?")
            If yadda = vbYes Then
                Chat.SendData "AVATAR" & "Mod" & vbCrLf
            End If
        'End If
    Next 'i
ElseIf yadda = "d" Then
    For i = 1 To 15
         If strMyUserName = strModName(i) And strModName(i) <> "" Then
            SuperModerator = True
            frmMod.Show
            frmMod.cmdKill.Visible = True
            frmMod.chkScrambler.Visible = True
            frmMod.cmdFreeze.Visible = True
            frmMod.cmdIP.Visible = True
            frmMod.cmdKick.Visible = True
            frmMod.cmdModWarn.Visible = True
            frmMod.cmdPraise.Visible = True
            frmMod.cmdGetPin.Visible = True
            frmMod.cmdPINBan.Visible = True
            frmMod.cmdBanIP.Visible = True
            cmd30Freeze.Visible = True
            yadda = MsgBox("Do you want to change your avatar to that of a moderator?", vbYesNo, "Change Avatar?")
            If yadda = vbYes Then
                Chat.SendData "AVATAR" & "SMod" & vbCrLf
            End If
        End If
    Next 'i
ElseIf yadda = "ÅEﬁﬁÅE" Then   ' strMyUserName = "dragoon" Or strMyUserName = "kumo" Then
        frmMod.Show
        Admin = True
        frmMod.cmdKill.Visible = True
        frmMod.cmdBan.Visible = True
        frmMod.cmdReset.Visible = True
        frmMod.chkScrambler.Visible = True
        frmMod.cmdBanIP.Visible = True
        frmMod.cmdFreeze.Visible = True
        frmMod.cmdIP.Visible = True
        frmMod.cmdKick.Visible = True
        frmMod.cmdModWarn.Visible = True
        frmMod.cmdPraise.Visible = True
        cmd30Freeze.Visible = True
        frmMod.cmdGetPin.Visible = True
        frmMod.cmdPINBan.Visible = True
        frmMod.cmdCloseGame.Visible = True
        yadda = MsgBox("Do you want to change your avatar to that of an admin?", vbYesNo, "Change Avatar?")
        If yadda = vbYes Then
            Chat.SendData "AVATAR" & "Admin" & vbCrLf
        End If
    End If
Else
    frmMod.Hide
    Admin = False
    Moderator = False
    SuperModerator = False
    frmMod.cmdKill.Visible = False
    frmMod.cmdBan.Visible = False
    frmMod.cmdReset.Visible = False
    frmMod.chkScrambler.Visible = False
    frmMod.cmdBanIP.Visible = False
    frmMod.cmdFreeze.Visible = False
    frmMod.cmdIP.Visible = False
    frmMod.cmdKick.Visible = False
    frmMod.cmdPraise.Visible = False
    frmMod.cmdModWarn.Visible = False
    cmd30Freeze.Visible = False
    frmMod.cmdGetPin.Visible = False
    frmMod.cmdPINBan.Visible = False
    Chat.SendData "NOTAWAY" & strMyUserName & vbCrLf
End If


End Sub
Private Sub SendPic()
On Error Resume Next
If strChar(1) = "" Then 'If my character doesn't exist
    Chat.SendData "NEEDCHAR" & vbCrLf
    Exit Sub
End If

Dim iCustChar As Long
iCustChar = FindWhichCharacter(strChar(1))

If strChar(1) = "Mia" Then
    IsaacM(curIsaac).Num = 0
End If
If strChar(1) = "Alex" Then
    IsaacM(curIsaac).Num = 1
End If
If strChar(1) = "Guard" Then
    IsaacM(curIsaac).Num = 2
End If
If strChar(1) = "Gladiator" Then
    IsaacM(curIsaac).Num = 3
End If
If strChar(1) = "Felix" Then
    IsaacM(curIsaac).Num = 4
End If
If strChar(1) = "Garret" Then
    IsaacM(curIsaac).Num = 5
End If
If strChar(1) = "Isaac" Then
    IsaacM(curIsaac).Num = 6
End If
If strChar(1) = "Ivan" Then
    IsaacM(curIsaac).Num = 7
End If
If strChar(1) = "Jenna" Then
    IsaacM(curIsaac).Num = 8
End If
If strChar(1) = "Kraden" Then
    IsaacM(curIsaac).Num = 9
End If
If strChar(1) = "Menardi" Then
    IsaacM(curIsaac).Num = 10
End If
If strChar(1) = "Saturos" Then
    IsaacM(curIsaac).Num = 11
End If
If strChar(1) = "Caption Contest Character" Then
    IsaacM(curIsaac).Num = 12
End If
If strChar(1) = "Piers" Then
    IsaacM(curIsaac).Num = 13
End If
If strChar(1) = "Kenny" Then
    IsaacM(curIsaac).Num = 14
End If
If strChar(1) = "KOS" Then
    IsaacM(curIsaac).Num = 16
End If
If strChar(1) = "Cloud" Then
    IsaacM(curIsaac).Num = 17
End If
If strChar(1) = "Sheba" Then
    IsaacM(curIsaac).Num = 19
End If
If strChar(1) = "Purple Piers" Then
    IsaacM(curIsaac).Num = 18
End If
If strChar(1) = "Agiato" Then
    IsaacM(curIsaac).Num = 20
End If
If strChar(1) = "Karst" Then
    IsaacM(curIsaac).Num = 21
End If
If strChar(1) = "The Wise One" Then
    IsaacM(curIsaac).Num = 22
End If
If strChar(1) = "Young Isaac" Then
    IsaacM(curIsaac).Num = 23
End If
If strChar(1) = "Young Garet" Then
    IsaacM(curIsaac).Num = 24
End If
If iCustChar <> 999 Then
    IsaacM(curIsaac).CustomCharacter = True
    Call AddCustChar(CInt(iCustChar))
    IsaacM(curIsaac).Num = picCustChar.UBound  'For now until mask/sprite implemented
Else
    IsaacM(curIsaac).CustomCharacter = False
End If

If IsaacM(curIsaac).CustomCharacter = False Then
    Chat.SendData "ISAACPIC" & IsaacM(curIsaac).Num & vbCrLf
Else
    Chat.SendData "ISAACPICC" & iCustChar & vbCrLf
End If

End Sub
Private Function CheckSwear() As Boolean
On Error Resume Next
Dim strCurse As String
Dim bSwear As Boolean
Dim strReturn As Long
bSwear = False
For i = 1 To 12
    If i = 1 Then
        strCurse = "fuck"
    End If
    If i = 2 Then
        strCurse = "shit"
    End If
    If i = 3 Then
        strCurse = "bitch"
    End If
    If i = 4 Then
        strCurse = "pussy"
    End If
    If i = 5 Then
        strCurse = "penis"
    End If
    If i = 6 Then
        strCurse = "nigger"
    End If
    If i = 7 Then
        strCurse = "asshole"
    End If
    If i = 8 Then
        strCurse = "dick"
    End If
    If i = 9 Then
        strCurse = "cunt"
    End If
    If i = 10 Then
        strCurse = "piss"
    End If
    If i = 11 Then
        strCurse = "gay"
    End If
    If i = 12 Then
        strCurse = "fag"
    End If
    Dim strTxt As String
    strTxt = LCase(txtmsg.Text)
    strReturn = InStr(1, strTxt, strCurse, vbTextCompare)
    If strReturn > 0 Then
        bSwear = True
    End If
Next 'i
If bSwear = True Then CheckSwear = True
If bSwear = False Then CheckSwear = False
End Function
Function GetRandomWord() As String
Dim intRand As Integer 'random integer
intRand = Int(Rnd * 10) + 1

Select Case intDifficulty
    Case 1 'Easy
        Select Case intRand
            Case 1
                GetRandomWord = "dog"
            Case 2
                GetRandomWord = "cat"
            Case 3
                GetRandomWord = "fish"
            Case 4
                GetRandomWord = "cow"
            Case 5
                GetRandomWord = "bat"
            Case 6
                GetRandomWord = "rat"
            Case 7
                GetRandomWord = "bird"
            Case 8
                GetRandomWord = "ape"
            Case 9
                GetRandomWord = "lice"
            Case Else
                GetRandomWord = "lion"
        End Select
    Case 2 'Medium
        Select Case intRand
            Case 1
                GetRandomWord = "Alex"
            Case 2
                GetRandomWord = "Mia"
            Case 3
                GetRandomWord = "Garet"
            Case 4
                GetRandomWord = "Ivan"
            Case 5
                GetRandomWord = "Babi"
            Case 6
                GetRandomWord = "Iodem"
            Case 7
                GetRandomWord = "Isaac"
            Case 8
                GetRandomWord = "Piers"
            Case 9
                GetRandomWord = "Kraden"
            Case Else
                GetRandomWord = "Jenna"
        End Select
    Case 3 'Hard
        Select Case intRand
            Case 1
                GetRandomWord = "Sol Blade"
            Case 2
                GetRandomWord = "Venus Djinn"
            Case 3
                GetRandomWord = "Froth Sphere"
            Case 4
                GetRandomWord = "Black Orb"
            Case 5
                GetRandomWord = "Cloak Ball"
            Case 6
                GetRandomWord = "High Impact"
            Case 7
                GetRandomWord = "Mars Stone"
            Case 8
                GetRandomWord = "Shaman's Rod"
            Case 9
                GetRandomWord = "Perfect Ply"
            Case Else
                GetRandomWord = "Lucky Medal"
        End Select
    Case 4 'Insane
        Select Case intRand
            Case 1
                GetRandomWord = "Golden Sun 2: The Lost Age"
            Case 2
                GetRandomWord = "You look like an Isaac to me."
            Case 3
                GetRandomWord = "Isaac gave Mia a Hard Nut."
            Case 4
                GetRandomWord = "I got a Kikuichimonji!"
            Case 5
                GetRandomWord = "Isaac did 243 damage."
            Case 6
                GetRandomWord = "Bring it on, whelp!"
            Case 7
                GetRandomWord = "Oh, that's right..."
            Case 8
                GetRandomWord = "I shall grant your wish..."
            Case 9
                GetRandomWord = "Who speaks to my mind?"
            Case Else
                GetRandomWord = "You made me spill my water!"
        End Select
End Select

End Function
Sub DetectScramble(ByVal strText As String)
'On Error Resume Next
bScramBusy = True

Dim intPointsAdded As Long
Dim strSWord
Dim strcurText As String
Dim strPlayer As String
'Dim strText As String
strSWord = Split(strText, ":", -1, vbTextCompare)

For i = 0 To UBound(strSWord)
    If i = 0 Then
        strPlayer = strSWord(i)
    End If
    If i = 1 Then
        strcurText = Mid$(strSWord(i), 2, Len(strSWord(i)))
        strcurText = LCase(strcurText)
    End If
Next 'i

Dim bNewSPlayer As Boolean
Dim FirstNewPlayer As Integer
Dim CorrectAnswer As Boolean
CorrectAnswer = False
bNewSPlayer = True
lstScores.Clear
For i = 1 To 20
    If SPlayers(i).Name <> "" Then
        lstScores.AddItem SPlayers(i).Name & ": " & SPlayers(i).Score
    End If
Next 'i

If strcurText = txtSOrgWord.Text Then
    CorrectAnswer = True
    For i = 1 To 20
        If TotalWinners < hSWinners.Value Then
            If strPlayer = SPlayers(i).Name Then
                If TotalWinners = 0 Then
                    If hSWinners.Value > 1 Then
                        SPlayers(i).Score = SPlayers(i).Score + (hSPoints.Value * 2)
                        intPointsAdded = hSPoints.Value * 2
                    Else
                        SPlayers(i).Score = SPlayers(i).Score + hSPoints.Value
                        intPointsAdded = hSPoints.Value
                    End If
                Else
                    SPlayers(i).Score = SPlayers(i).Score + hSPoints.Value
                End If
                bNewSPlayer = False 'Not a new player
    
                GoTo subCorrect
                Exit Sub
                DoEvents
            End If
        End If

        DoEvents
    Next 'i

End If

If bNewSPlayer = True And CorrectAnswer = True Then
    FirstNewPlayer = 999
    For i = 1 To 20
        If TotalWinners < hSWinners.Value Then
            If FirstNewPlayer = 999 Then
                If SPlayers(i).Name = "" Then
                    FirstNewPlayer = i
                    SPlayers(FirstNewPlayer).Name = strPlayer 'Create player
                    If TotalWinners = 0 And hSWinners.Value <> 1 Then
                        SPlayers(FirstNewPlayer).Score = SPlayers(FirstNewPlayer).Score + (hSPoints.Value * 2)
                        intPointsAdded = hSPoints.Value * 2
                    Else
                        SPlayers(FirstNewPlayer).Score = SPlayers(FirstNewPlayer).Score + hSPoints.Value
                        intPointsAdded = hSPoints.Value
                    End If
                    GoTo subCorrect
                    Exit Sub
                End If
                DoEvents
            End If
        End If
        DoEvents
    Next 'i

End If

bScramBusy = False

Exit Sub
subCorrect:
If timeSWait.Enabled = False Then
    'If hSWinners.Value <> 1 And TotalWinners = 0 Then
        'Chat.SendData "GAMETXT" & strMacro(0) & strPlayer & " got the correct answer (" & txtSOrgWord.Text & "). " & strPlayer & " won " & CStr(hSPoints.Value * 2) & " points.  No more winners remaining." & strMacro(1) & vbCrLf
        'DoEvents
    'Else
        'Chat.SendData "GAMETXT" & strMacro(0) & strPlayer & " got the correct answer (" & txtSOrgWord.Text & "). " & strPlayer & " won " & CStr(hSPoints.Value) & " points." & (hSWinners.Value - TotalWinners) & " winners remaining." & strMacro(1) & vbCrLf
        'DoEvents
    'End If
    
    'DoEvents
    TotalWinners = TotalWinners + 1
    
    If TotalWinners <> hSWinners.Value Then
        
        Chat.SendData "METXT" & strMacro(0) & strPlayer & " got the correct answer (" & txtSOrgWord.Text & "). " & strPlayer & " won " & CStr(intPointsAdded) & " points." & (hSWinners.Value - TotalWinners) & " winners remaining." & strMacro(1) & vbCrLf
        DoEvents
        Exit Sub
    Else
        Chat.SendData "METXT" & strMacro(0) & strPlayer & " got the correct answer (" & txtSOrgWord.Text & "). " & strPlayer & " won " & CStr(intPointsAdded) & " points.  No more winners remaining." & strMacro(1) & vbCrLf
        DoEvents
        timeSWait.Enabled = True
        hSPoints.Enabled = True
        timeAnswer.Enabled = False
        iAnswers = 0
        Exit Sub
    End If

End If
    bScramBusy = False
End Sub

Sub StartScramble()
On Error Resume Next
TotalWinners = 0
txtSOrgWord.Text = LCase(txtSOrgWord.Text)
Chat.SendData "GAMEMETXT" & strMacro(0) & "Unscramble:     " & txtSScrambled.Text & "     In the category: " & txtSCategory.Text & " for " & CStr(hSPoints.Value) & " points with " & hSWinners.Value & " winners." & strMacro(1) & vbCrLf
DoEvents
'Chat.SendData "GAMETXT" & strMacro(0) & f & strMacro(1) & vbCrLf
cmdScramble.Caption = "Stop"
If chkAnswers.Value = 1 Then
    iAnswers = 0
    timeAnswer.Enabled = True
End If

End Sub
Function CheckCaps(strMsg As String) As Boolean
On Error Resume Next
Dim CapsChar As Boolean
Dim LowerChar As Boolean
'If InStr(0, strMsg, "         ", vbTextCompare) = 0 Then
'    CheckCaps = True
'    Exit Function
'End If

CapsChar = False
LowerChar = False
Dim strCheck As Integer
For i = 1 To Len(strMsg)
    strCheck = Asc(Mid$(strMsg, i, 1))
    If strCheck >= 97 And strCheck <= 122 Then
        LowerChar = True
    End If
    If strCheck >= 65 And strCheck <= 90 Then
        CapsChar = True
    End If
Next 'i
If CapsChar = True And LowerChar = False Then
    CheckCaps = True
Else
    CheckCaps = False
End If
End Function
Sub AddAllChat(strChatTxt As String)
frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
frmChat.txtChat.SelLength = Len(strChatTxt)
frmChat.txtChat.SelBold = True
frmChat.txtChat.SelColor = &HC000&
frmChat.txtChat.SelFontName = "Arial"
frmChat.txtChat.SelText = vbNewLine & strChatTxt

frmChat.txtGame.SelStart = Len(frmChat.txtGame.Text)
frmChat.txtGame.SelLength = Len(strChatTxt)
frmChat.txtGame.SelBold = True
frmChat.txtGame.SelColor = &HC000&
frmChat.txtGame.SelFontName = "Verdana Ref"
frmChat.txtGame.SelText = vbNewLine & strChatTxt

'frmChat.txtTalk.SelStart = Len(frmChat.txtTalk.Text)
'frmChat.txtTalk.SelLength = Len(strChatTxt)
'frmChat.txtTalk.SelBold = True
'frmChat.txtTalk.SelColor = &HC000&
'frmChat.txtTalk.SelFontName = "Verdana Ref"
'frmChat.txtTalk.SelText = vbNewLine & strChatTxt
End Sub
Sub AddCustChar(iCustChar As Integer)
On Error Resume Next
Load picCustChar(picCustChar.UBound + 1)
    Load picCustCharM(picCustCharM.UBound + 1)
    Dim intNewPicture As Long
    intNewPicture = picCustChar.UBound
    picCustChar(intNewPicture).Cls
    picCustChar(picCustChar.UBound) = LoadPicture(App.Path & "\files\" & CustomChar(iCustChar).Picture & "S.gif")
    picCustChar(intNewPicture).Refresh
    'frmMultiplayer.picCustChar(intNewPicture).Visible = True
    
    picCustCharM(intNewPicture).Cls
    picCustCharM(picCustChar.UBound) = LoadPicture(App.Path & "\files\" & CustomChar(iCustChar).Picture & "M.gif")
    picCustCharM(intNewPicture).Refresh

End Sub
Sub GetProfile(strGetUser As String)
If strGetUser = strMyUserName Or Moderator = True Or Admin = True Then
    txtPName.Locked = False
    txtPAge.Locked = False
    txtPSex.Locked = False
    txtPLocation.Locked = False
    txtPAIM.Locked = False
    txtPMSN.Locked = False
    txtPEmail.Locked = False
    txtPOther.Locked = False
    cmdPSave.Enabled = True
    txtPICQ.Locked = False
    txtPURL.Locked = False
    cmdPAvatar.Enabled = True
Else
    txtPName.Locked = True
    txtPAge.Locked = True
    txtPSex.Locked = True
    txtPLocation.Locked = True
    txtPAIM.Locked = True
    txtPMSN.Locked = True
    txtPEmail.Locked = True
    txtPOther.Locked = True
    cmdPSave.Enabled = False
    txtPICQ.Locked = True
    txtPURL.Locked = True
    cmdPAvatar.Enabled = False
End If
If strGetUser <> "" Then
    Chat.SendData "GETPROFILE" & strGetUser & vbCrLf
    framProfile.Visible = True
    framProfile.Top = 248
    framProfile.Height = 209
End If
End Sub

Private Sub txtTalk_Change()
'Call AutoScroll(txtTalk)
End Sub
Function CheckIgnore(strUser As String) As Boolean
'Checks if a user is ignored
On Error Resume Next
If strUser <> "" Then
    For i = 0 To lstIgnore.ListCount
        If strUser = lstIgnore.List(i) Then
            CheckIgnore = True
            Exit Function
        End If
    Next 'i
End If
'Did not find the user to ignore
CheckIgnore = False
End Function
Sub UpdateUserDisplay()
On Error Resume Next
Dim ActualUser(1 To 20) As Long
Dim curActualUser As Long
Dim curLabel As Long
curActualUser = 0
'For i = 1 To 20
'    If Users(i).Enabled = True Then
'        curActualUser = curActualUser + 1
'        ActualUser(curActualUser) = i
'    End If
'Next 'i
curLabel = 0
vchat.Visible = True
For i = MinDisplayPlayer To MinDisplayPlayer + 12
    If i < lstUsers.ListCount - MinDisplayPlayer + 1 And lstUsers.List(i) <> "" Then
        lblUser(curLabel).Caption = lstUsers.List(i)
        Dim intTempUser As Long
        intTempUser = 0
        For q = 0 To 20
            If Users(q).Name = lstUsers.List(i) Then
                intTempUser = q
                Exit For
            End If
        Next 'q
        If Users(intTempUser).Admin = True Then
            lblUser(curLabel).ForeColor = RGB(255, 0, 0)
        ElseIf Users(intTempUser).Moderator = True Then
            lblUser(curLabel).ForeColor = RGB(0, 0, 255)
        End If
        
        If Users(intTempUser).Ignore = True Then
            imgIcon(curLabel).Picture = imgSAvatar(2).Picture
        ElseIf Users(intTempUser).Avatar = "Away" Then
            imgIcon(curLabel).Picture = imgSAvatar(1).Picture
        ElseIf Users(intTempUser).Avatar = "Admin" Then
            imgIcon(curLabel).Picture = imgSAvatar(0).Picture
        ElseIf Users(intTempUser).Avatar = "Mod" Then
            imgIcon(curLabel).Picture = imgSAvatar(3).Picture
        ElseIf Users(intTempUser).Avatar = "SMod" Then
            imgIcon(curLabel).Picture = imgSAvatar(4).Picture
        ElseIf Left$(Users(intTempUser).Avatar, 4) = "CUST" Then
            imgIcon(curLabel).Picture = imgSAvatar(CInt(Right$(Users(intTempUser).Avatar, 2))).Picture
        Else
            imgIcon(curLabel).Picture = LoadPicture(App.Path & "\icons\" & Users(intTempUser).Avatar & ".gif")
        End If
        
        'If Users(intTempUser).Avatar <> "Away" And Users(intTempUser).Ignore = False Then
        '    imgIcon(curLabel).Picture = LoadPicture(App.Path & "\icons\" & Users(intTempUser).Avatar & ".gif")
        'ElseIf Users(intTempUser).Ignore = True Then
        '    imgIcon(curLabel).Picture = imgSAvatar(2).Picture
        'ElseIf Users(intTempUser).Avatar = "Away" Then
        '    imgIcon(curLabel).Picture = imgSAvatar(1).Picture
        'End If
        
        Debug.Print Users(intTempUser).Name & " " & Users(intTempUser).Avatar
        lblUser(curLabel).Visible = True
        imgIcon(curLabel).Visible = True
        curLabel = curLabel + 1
        vchat.Max = lstUsers.ListCount - 13
    Else
        lblUser(curLabel).Visible = False
        imgIcon(curLabel).Visible = False
        vchat.Visible = False
        curLabel = curLabel + 1
    End If
Next 'i

If lstUsers.ListCount > 12 Then
    vchat.Visible = True
End If

End Sub

Private Sub vUser_Change()
MinDisplayPlayer = vchat.Value
Call UpdateUserDisplay
End Sub
