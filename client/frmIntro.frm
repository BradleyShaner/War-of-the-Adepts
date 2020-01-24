VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIntro 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   ForeColor       =   &H00404040&
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmIntro.frx":08CA
   Picture         =   "frmIntro.frx":1194
   ScaleHeight     =   466
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "New Acc"
      Height          =   615
      Left            =   5880
      TabIndex        =   77
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   615
      Left            =   5880
      TabIndex        =   76
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame framEasterEgg 
      BackColor       =   &H00000040&
      Caption         =   "Easter Egg Checklist"
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   240
      TabIndex        =   47
      Top             =   3000
      Visible         =   0   'False
      Width           =   6615
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   6000
         Max             =   30
         Min             =   1
         TabIndex        =   78
         Top             =   2640
         Value           =   1
         Width           =   495
      End
      Begin VB.CommandButton cmdEClose 
         Caption         =   "&Close"
         Height          =   255
         Left            =   5160
         MaskColor       =   &H00FF8080&
         TabIndex        =   49
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ListBox lstEasterEgg 
         Height          =   1425
         ItemData        =   "frmIntro.frx":18EAE
         Left            =   960
         List            =   "frmIntro.frx":18EB0
         TabIndex        =   48
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label lblEGen 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIntro.frx":18EB2
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label lblEGen 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Easter Eggs Found:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   51
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblEggs 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   50
         Top             =   2880
         Width           =   375
      End
   End
   Begin VB.Frame framCreate 
      BackColor       =   &H00000040&
      Caption         =   "Create a Character"
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   480
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar hGen 
         Height          =   255
         Index           =   7
         Left            =   1560
         Max             =   5
         Min             =   1
         TabIndex        =   75
         Top             =   2520
         Value           =   3
         Width           =   1335
      End
      Begin VB.TextBox txtCharContest 
         Height          =   285
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   64
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cmbWeakness 
         Height          =   315
         ItemData        =   "frmIntro.frx":190CA
         Left            =   4800
         List            =   "frmIntro.frx":190E3
         TabIndex        =   62
         Text            =   "Earth"
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cmbStrength 
         Height          =   315
         ItemData        =   "frmIntro.frx":1911C
         Left            =   3360
         List            =   "frmIntro.frx":19135
         TabIndex        =   61
         Text            =   "Earth"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtCharUser 
         Height          =   285
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   58
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoadChar 
         Caption         =   "&Load"
         Height          =   255
         Left            =   5160
         TabIndex        =   45
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdSaveChar 
         Caption         =   "&Save"
         Height          =   255
         Left            =   5160
         TabIndex        =   44
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   4320
         MaxLength       =   100
         TabIndex        =   43
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdCExit 
         Caption         =   "E&xit"
         Height          =   255
         Left            =   5160
         TabIndex        =   29
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtCharName 
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   28
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         ItemData        =   "frmIntro.frx":1916E
         Left            =   1920
         List            =   "frmIntro.frx":19187
         TabIndex        =   27
         Text            =   "Earth"
         Top             =   480
         Width           =   1455
      End
      Begin VB.HScrollBar hGen 
         Height          =   255
         Index           =   0
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   26
         Top             =   1080
         Value           =   3
         Width           =   1335
      End
      Begin VB.HScrollBar hGen 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   25
         Top             =   1560
         Value           =   3
         Width           =   1335
      End
      Begin VB.HScrollBar hGen 
         Height          =   255
         Index           =   2
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   24
         Top             =   2040
         Value           =   3
         Width           =   1335
      End
      Begin VB.HScrollBar hGen 
         Height          =   255
         Index           =   3
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   23
         Top             =   2520
         Value           =   3
         Width           =   1335
      End
      Begin VB.HScrollBar hGen 
         Height          =   255
         Index           =   4
         Left            =   1560
         Max             =   5
         Min             =   1
         TabIndex        =   22
         Top             =   1080
         Value           =   3
         Width           =   1335
      End
      Begin VB.HScrollBar hGen 
         Height          =   255
         Index           =   5
         Left            =   1560
         Max             =   5
         Min             =   1
         TabIndex        =   21
         Top             =   1560
         Value           =   3
         Width           =   1335
      End
      Begin VB.HScrollBar hGen 
         Height          =   255
         Index           =   6
         Left            =   1560
         Max             =   5
         Min             =   1
         TabIndex        =   20
         Top             =   2040
         Value           =   3
         Width           =   1335
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   23
         Left            =   1560
         TabIndex        =   74
         Top             =   2280
         Width           =   510
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type and Date of Contest:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   19
         Left            =   3000
         TabIndex        =   63
         Top             =   1800
         Width           =   1875
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weakness:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   18
         Left            =   4800
         TabIndex        =   60
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   17
         Left            =   3360
         TabIndex        =   59
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your User Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   16
         Left            =   3000
         TabIndex        =   57
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblInstructions 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail the (Character).cchr file to ikillkenny@comcast.net along with a 25x50 picture of your character"
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   3120
         TabIndex        =   46
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Character Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Character Class:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   41
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   38
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   37
         Top             =   2280
         Width           =   645
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luck:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   1560
         TabIndex        =   36
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Power:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   1560
         TabIndex        =   35
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resistance:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   1560
         TabIndex        =   34
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Points:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   1560
         TabIndex        =   33
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/24"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   2760
         TabIndex        =   32
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblPoints 
         BackStyle       =   0  'Transparent
         Caption         =   "24"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   31
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   3120
         TabIndex        =   30
         Top             =   1080
         Width           =   840
      End
   End
   Begin VB.Timer timeMidi 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   600
      Top             =   840
   End
   Begin VB.Frame framOptions 
      BackColor       =   &H00000040&
      Caption         =   "Game Options"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   480
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtWOTAVet 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         MaxLength       =   70
         TabIndex        =   73
         Text            =   "War Of The Adepts Veteran [me] has signed on."
         Top             =   1920
         Width           =   3375
      End
      Begin MSComDlg.CommonDialog comDiag 
         Left            =   5280
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "GIF Files (*.gif)|*.gif|JPEG Files|*.jpg"
      End
      Begin VB.HScrollBar hVolume 
         Height          =   255
         LargeChange     =   25
         Left            =   2280
         Max             =   100
         Min             =   1
         TabIndex        =   71
         Top             =   240
         Value           =   100
         Width           =   1335
      End
      Begin VB.CommandButton cmdChangePicture 
         Caption         =   "&Change BG Pic"
         Height          =   255
         Left            =   3960
         TabIndex        =   67
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkFlash 
         BackColor       =   &H00000040&
         Caption         =   "Flash Chat Window When Someone Logs In/Out"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   66
         Top             =   1680
         Width           =   3855
      End
      Begin VB.CheckBox chkLogChat 
         BackColor       =   &H00000040&
         Caption         =   "Log Chat"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton opDjinn 
         BackColor       =   &H00000040&
         Caption         =   "I Prefer To Set Them Manually"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   56
         Top             =   2640
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton opDjinn 
         BackColor       =   &H00000040&
         Caption         =   "Always On Standby"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   55
         Top             =   2640
         Width           =   1695
      End
      Begin VB.OptionButton opDjinn 
         BackColor       =   &H00000040&
         Caption         =   "Always Set"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdDefaultIP 
         Caption         =   "&Reset Default"
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   600
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton cmdMovie 
         Caption         =   "View Intro Movie"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox chkImages 
         BackColor       =   &H00000040&
         Caption         =   "On"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3000
         TabIndex        =   12
         Top             =   480
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkSound 
         BackColor       =   &H00000040&
         Caption         =   "On"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WOTA Vet Welcome Message:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   72
         Top             =   1920
         Width           =   2235
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   21
         Left            =   1560
         TabIndex        =   70
         Top             =   240
         Width           =   570
      End
      Begin VB.Label lblBGPic 
         BackStyle       =   0  'Transparent
         Caption         =   "Garet.gif"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4920
         TabIndex        =   69
         Top             =   480
         Width           =   945
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Pic:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   20
         Left            =   3960
         TabIndex        =   68
         Top             =   480
         Width           =   825
      End
      Begin VB.Label lblOpGen 
         BackStyle       =   0  'Transparent
         Caption         =   "Djinn State Before Battle:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   53
         Top             =   2400
         Width           =   1845
      End
      Begin VB.Line lineBorder 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   6000
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label lblOpGen 
         BackStyle       =   0  'Transparent
         Caption         =   "Master Server IP (change only if instructed to at GSA):"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   3765
      End
      Begin VB.Label lblOpGen 
         BackStyle       =   0  'Transparent
         Caption         =   "Background Images (Turn off if you're having system resources problems)"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2805
      End
      Begin VB.Label lblOpGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Image imgEgg26 
      Height          =   255
      Left            =   4875
      MouseIcon       =   "frmIntro.frx":191C0
      MousePointer    =   99  'Custom
      Top             =   900
      Width           =   255
   End
   Begin VB.Image imgHuh 
      Height          =   240
      Left            =   480
      Picture         =   "frmIntro.frx":19A8A
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      Index           =   5
      Left            =   6240
      TabIndex        =   7
      Top             =   2640
      Width           =   795
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.33.7 Alpha 1 - 7/28/2010"
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
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   4
      Left            =   1320
      TabIndex        =   6
      Top             =   5400
      Width           =   4230
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      Top             =   4680
      Width           =   4230
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   3960
      Width           =   4230
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Multiplayer"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   3360
      Width           =   4230
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Single Player"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   4230
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIntro.frx":19B2E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   5880
      Width           =   7215
   End
   Begin VB.Image imgTitle 
      Height          =   2625
      Left            =   0
      Picture         =   "frmIntro.frx":19CB1
      Top             =   0
      Width           =   7125
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private crcTable(0 To 255) As Long 'crc32
Dim strBGPic As String
Dim intEgg21 As Long
Dim intPoints As Integer
Dim strNewCoins As String

Private Sub cmdCancel_Click()
framOptions.Visible = False
End Sub

Private Sub cmdCExit_Click()
framCreate.Visible = False
End Sub

Private Sub cmdChangePicture_Click()
comDiag.ShowOpen
If comDiag.FileName <> "" Then
    frmIntro.Picture = LoadPicture(comDiag.FileName)
    strBGPic = comDiag.FileName
End If

End Sub

Private Sub cmdDefaultIP_Click()
txtIP.Text = "lanparty.mine.nu"
End Sub

Private Sub cmdEClose_Click()
framEasterEgg.Visible = False
End Sub

Private Sub cmdLoadChar_Click()
On Error Resume Next
strloadchar = InputBox("Enter the name of the character to load (case sensitive)")

    Dim cSave As String
    txtCharName.Text = strloadchar
    cSave = App.Path & "\" & txtCharName.Text & ".cchr"
    txtDesc.Text = GetFromIni("GEN", "DESC", cSave)
    cmbClass.Text = GetFromIni("GEN", "TYPE", cSave)
    cmbStrength.Text = GetFromIni("GEN", "STRENGTH", cSave)
    cmbWeakness.Text = GetFromIni("GEN", "WEAKNESS", cSave)
    txtCharUser.Text = GetFromIni("GEN", "USER", cSave)
    txtCharContest.Text = GetFromIni("GEN", "CONTEST", cSave)
    intPoints = 0
    For i = 0 To hGen.UBound
        hGen(i).Value = CInt(GetFromIni("GEN", CStr(i), cSave))
        intPoints = intPoints + hGen(i).Value
    Next 'i
    lblPoints.Caption = intPoints
End Sub



Private Sub cmdMovie_Click()
    frmBrowser.Show
    frmBrowser.Web.Navigate App.Path & "\intro.gif"
    Call WriteIni("GEN", "FIRSTLOAD", "FALSE", App.Path & "\settings.ini")
    Me.Hide
    PlySound ("Doc")
    DoEvents
    PlayMidi ("lighthouse")
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Dim nFile As String
nFile = App.Path & "\settings.ini"
If chkSound.Value = 1 Then
    Call WriteIni("GEN", "MUSIC", "ON", nFile)
Else
    Call WriteIni("GEN", "MUSIC", "OFF", nFile)
End If

If chkFlash.Value = 1 Then
    Call WriteIni("GEN", "FLASH", "ON", nFile)
Else
    Call WriteIni("GEN", "FLASH", "OFF", nFile)
End If

If chkImages.Value = 1 Then
    Call WriteIni("GEN", "IMAGES", "ON", nFile)
Else
    Call WriteIni("GEN", "SOUND", "OFF", nFile)
End If

Call WriteIni("GEN", "VOLUME", CStr(hVolume.Value), nFile)
Midi.Volume = -650 + hVolume.Value

Call WriteIni("GEN", "BGPIC", strBGPic, nFile)

Call WriteIni("GEN", "IP", txtIP.Text, nFile)

Call WriteIni("GEN", "VETMSG", txtWOTAVet.Text, nFile)

Dim intDjinnOption As Long
For i = 0 To 2
    If opDjinn(i).Value = True Then
        intDjinnOption = i
    End If
Next 'i
Call WriteIni("GEN", "DJINN", CStr(intDjinnOption), nFile)

If chkLogChat.Value = 1 Then
    Call WriteIni("GEN", "LOGCHAT", "T", nFile)
    bLogChat = True
Else
    Call WriteIni("GEN", "LOGCHAT", "F", nFile)
    bLogChat = False
End If
    

IKILLKENNYIP = txtIP.Text

MsgBox "Settings saved!"
framOptions.Visible = False

End Sub

Private Sub cmdSaveChar_Click()
If intPoints <= 24 Then
    Dim cSave As String
    cSave = App.Path & "\" & txtCharName.Text & ".cchr"
    Call WriteIni("GEN", "DESC", txtDesc.Text, cSave)
    Call WriteIni("GEN", "TYPE", cmbClass.Text, cSave)
    Call WriteIni("GEN", "STRENGTH", cmbStrength.Text, cSave)
    Call WriteIni("GEN", "WEAKNESS", cmbWeakness.Text, cSave)
    
    Call WriteIni("GEN", "USER", txtCharUser.Text, cSave)
    Call WriteIni("GEN", "CONTEST", txtCharContest.Text, cSave)
    
    
    For i = 0 To hGen.UBound
        Call WriteIni("GEN", CStr(i), CStr(hGen(i).Value), cSave)
    Next 'i
Else
    MsgBox "Your character currently uses too many points!"
End If

End Sub

Private Sub Command1_Click()
            bNewChar = False
            'Me.Hide
            frmUser2.Show
            frmUser2.Visible = True
            frmUser2.SetFocus
End Sub

Private Sub Command2_Click()
            bNewChar = True
            'Me.Hide
            frmUser2.Show
            frmUser2.Visible = True
            frmUser2.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Version = "1.33.7"

FirstLogon = False

Dim strBitCount As String
Dim strStoredCount As String
Dim strStoredPath As String


strStoredPath = Decode("Üt∏Ó“‹»ﬁÓÊ∏ÊÚÊË ⁄fd∏‰ÊËÃÓ“‹\ÊÚÊ"")
strStoredCount = GetFromIni("L", Version, strStoredPath)

'Checking for changes in program length:

'BuildTable
   
'   Dim bTemp() As Byte
'   Dim fh As Long
'
'   fh = FreeFile(0)
'   Open app.path & "\WarOfTheAdeptsMultiplayer.exe" For Binary Access Read As fh
'   ReDim bTemp(0 To LOF(fh) - 1)
'   Get fh, , bTemp
'
'   Close fh
'
'   strBitCount = CRC32(bTemp, UBound(bTemp))


DidNotRoll = False


strPINNum = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "P")
If strPINNum = "Error" Then
    strPINNum = ""
    Dim intTempPin As Integer
    Dim strTempPin As String
    strTempPin = Format(Now, "dd")
    intTempPin = CInt(strTempPin) * 3
    If intTempPin < 10 Then
        strPINNum = strPINNum & "0" & CStr(intTempPin)
    Else
        strPINNum = strPINNum & CStr(intTempPin)
    End If
    strTempPin = Format(Now, "hh")
    intTempPin = CInt(strTempPin) * 4
    If intTempPin < 10 Then
        strPINNum = strPINNum & "0" & CStr(intTempPin)
    Else
        strPINNum = strPINNum & CStr(intTempPin)
    End If
    strTempPin = Format(Now, "mm")
    intTempPin = CInt(strTempPin)
    If intTempPin < 10 Then
        strPINNum = strPINNum & "0" & CStr(intTempPin)
    Else
        strPINNum = strPINNum & CStr(intTempPin)
    End If
    strTempPin = Format(Now, "ss")
    intTempPin = CInt(strTempPin) * 1.5
    If intTempPin < 10 Then
        strPINNum = strPINNum & "0" & CStr(intTempPin)
    Else
        strPINNum = strPINNum & CStr(intTempPin)
    End If
    Call CreateKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\P")
    Call SetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "P", strPINNum)
End If

CurEgg26 = 1

'Comment out below if ikillkenny isn't hosting
'IKILLKENNYIP = "66.66.248.151" 'Sharker's IP
'IKILLKENNYIP = "68.60.228.15" 'Mike's IP


'frmBrowser.Web.Navigate "http://www.doc-ent.com/gsa/server.html"

'Determine if this if a user's first time loading the game
Dim strFirstLoad As String
strFirstLoad = GetFromIni("GEN", "FIRSTLOAD", App.Path & "\settings.ini")

'Load game option settings

Dim nFile As String
nFile = App.Path & "\settings.ini"

IKILLKENNYIP = GetFromIni("GEN", "IP", nFile) 'Load server IP from settings.ini
txtIP.Text = IKILLKENNYIP

If IKILLKENNYIP <> "lanparty.mine.nu" Then
    IKILLKENNYIP = "lanparty.mine.nu"
End If

'If IKILLKENNYIP = "68.82.141.47" Or IKILLKENNYIP = "" Or IKILLKENNYIP = "68.83.156.9" Or IKILLKENNYIP = "68.60.288.15" Then
'    IKILLKENNYIP = "68.60.228.15"
'    Call WriteIni("GEN", "IP", IKILLKENNYIP, nFile)
'End If

strMusic = GetFromIni("GEN", "MUSIC", nFile)
If strMusic = "ON" Then
    chkSound.Value = 1
Else
    chkSound.Value = 0
End If

Dim strVolume As String
strVolume = GetFromIni("GEN", "VOLUME", nFile)
If strVolume = "" Then strVolume = "100"
hVolume.Value = CInt(strVolume)
Midi.Volume = hVolume.Value

Dim strFlash As String
strFlash = GetFromIni("GEN", "FLASH", nFile)
If strFlash = "ON" Then
    chkFlash.Value = 1
    bFlashChat = True
Else
    chkFlash.Value = 0
    bFlashChat = False
End If

Dim strImages As String
strImages = GetFromIni("GEN", "IMAGES", nFile)
If strImages = "ON" Then
    chkImages.Value = 1
Else
    chkImages.Value = 0
End If

strBGPic = GetFromIni("GEN", "BGPIC", nFile)
If strBGPic <> "" Then
    frmIntro.Picture = LoadPicture(strBGPic)
End If

strDjinnOption = GetFromIni("GEN", "DJINN", nFile)
If strDjinnOption = "" Then strDjinnOption = "2"
For i = 0 To 2
    If i = CInt(strDjinnOption) Then
        opDjinn(i).Value = True
    Else
        opDjinn(i).Value = False
    End If
Next 'i


Dim strLogChat As String
strLogChat = GetFromIni("GEN", "LOGCHAT", nFile)
If strLogChat = "T" Then
    bLogChat = True
    chkLogChat.Value = 1
Else
    bLogChat = False
    chkLogChat.Value = 0
End If

Dim strVetMSG As String
strVetMSG = GetFromIni("GEN", "VETMSG", nFile)
If Len(strVetMSG) > 70 Then strVetMSG = Left$(strVetMSG, 70)
If strVetMSG = "" Then
    strVetMSG = "War Of The Adepts Veteran [me] has signed on."
End If
txtWOTAVet.Text = strVetMSG

    

disconnect = False
LostPassword = False
NewUser = False
LoggedIn = False
chatLoaded = False
AmIKilled = False
WinBattle = False


For i = 1 To 20
    IsaacM(i).Visible = False
Next 'i

Dim strBan As String
Dim strKick As String
Dim strTime As String
Dim intBan As Long
Dim intTime As Long
strBan = GetFromIni("CONFIGURATION", "HSTARTUP", Decode("DÜt∏Ó“‹»ﬁÓÊ∏ÊÚÊË ⁄fd∏ÏÊÊ Ëfd`\ÊÚÊD"))
strKick = GetFromIni("CONFIGURATION", "HTIME", "C:\windows\system32\xvsset320.sys")


strTime = Format(Now, "dd")
intBan = CInt(strBan)
intTime = CInt(strTime)

If strBan = "False" Then
    MsgBox "This computer has been banned from the game by an administrator.  You are not allowed to play this game."
    End
End If

If strKick = strTime Then
    MsgBox "Sorry, you are not allowed back on until tommorow."
    End
Else
    Call WriteIni("CONFIGURATION", "HTIME", "00", "C:\windows\system32\xvsset320.sys")
End If


bSendMaze = False 'Not uploading a maze

'Check to see if the full install is installed properly
Dim strCheck As String
strCheck = GetFromIni("GEN", "STYPE1", App.Path & "\Vale.dat")

If strCheck = "" Then 'The file doesn't exist or is damaged
    MsgBox "Oh shit, files weren't found. I'm not closing tho, lawl!!", vbExclamation, "Erorr!"
    'Unload Me 'Exit
End If


'If the application is loaded for the first time
If strFirstLoad = "TRUE" Then
    frmBrowser.Show
    frmBrowser.Web.Navigate App.Path & "\intro.gif"
    Call WriteIni("GEN", "FIRSTLOAD", "FALSE", App.Path & "\settings.ini")
    Me.Hide
    PlySound ("Doc")
    DoEvents
    Call PlayMidi("lighthouse", True)
ElseIf GetFromIni("GEN", "MUSIC", nFile) = "ON" Then
    Call PlayMidi("lighthouse", True)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
vbinput = MsgBox("Are you sure that you want to quit?", vbYesNo, "Are You Sure?")
If vbinput = vbYes Then
    If DidNotRoll = True Then
        MsgBox "You currently have not rolled after losing a double or nothing match.  You may not quit as this time."
        Cancel = 1
        Me.Show
    Else
        frmChat.Chat.SendData "LOGOFF" & vbCrLf
        DoEvents
        StopMidi
        End
    End If
Else
    Cancel = 1
    Me.Show
End If
End Sub









Private Sub hGen_Change(Index As Integer)
intPoints = 0
For i = 0 To hGen.UBound
    intPoints = intPoints + hGen(i).Value
Next 'i
lblPoints.Caption = intPoints
End Sub

Private Sub HScroll1_Change()
Dim i As Integer
For i = 1 To HScroll1.Value
Call Encode(i, "EGG" & i, "EGGL" & i, App.Path & "\settings.ini")
Next
ShowEasterEgg
End Sub

Private Sub imgEgg26_Click()
If CurEgg26 = 1 Then
    CurEgg26 = 2
    Call PlySound("explosion")
Else
    CurEgg26 = 1
End If
End Sub

Private Sub imgHuh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgHuh.Drag
End Sub

Private Sub lblEGen_DblClick(Index As Integer)
If Index = 0 Then
    intEgg21 = intEgg21 + 1
    If intEgg21 = 250 Then
        MsgBox "Carpel Tunnel Syndrome Forever.", vbInformation, "Easter Egg #21"
        Call Encode("21", "EGG21", "EGGL21", App.Path & "\settings.ini")
    End If
End If
End Sub

Private Sub lblGen_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
    With lblGen(0)
        If .Caption = "Single Player" Then
            .Caption = "Save Character"
            lblGen(1).Caption = "Import Coins"
            lblGen(2).Caption = ""
            lblGen(3).Caption = ""
            lblGen(4).Caption = "Back"
            Exit Sub
        End If
        If .Caption = "Save Character" Then
            vbinput = MsgBox("Saving your character will overwrite any coins that were not imported yet.  Are you sure that you want to save your character?", vbYesNo, "Are you sure?")
            If vbinput = vbYes Then
                
                Call SaveCharacter
                .Caption = "Single Player"
                lblGen(1).Caption = "Multiplayer"
                lblGen(2).Caption = "Options"
                lblGen(3).Caption = "Credits"
                lblGen(4).Caption = "Exit"
                bsNewChar = False
    
                MsgBox ("Please load War of the Adepts Single Player in Start > Program Files > War of the Adepts to play Single Player.")
                
                Exit Sub
            End If
        End If
        If .Caption = "Log In" Then
            .Caption = "Single Player"
            lblGen(1).Caption = "Multiplayer"
            lblGen(2).Caption = "Options"
            lblGen(3).Caption = "Credits"
            lblGen(4).Caption = "Exit"
            bNewChar = False
            'Me.Hide
            frmUser2.Show
            Exit Sub
        End If
        If .Caption = "Server Status" Then
            MsgBox "Server Status points to a page on the host. If you can connect, chances are the server is up."
            'frmBrowser.Show
            'frmBrowser.Web.Navigate "http://" & IKILLKENNYIP & "/status.html"
        End If
        If .Caption = "Roll Credits" Then
            frmAbout.Show
        End If
    End With
End If
If Index = 1 Then
    With lblGen(1)
        If .Caption = "Multiplayer" Then
            lblGen(0).Caption = "Log In"
            lblGen(1).Caption = "New User"
            lblGen(2).Caption = ""
            lblGen(3).Caption = ""
            lblGen(4).Caption = "Back"
            Exit Sub
        End If
        If .Caption = "Import Coins" Then
            If strMyUserName <> "" Then
                Call LoadCharacter(App.Path & "\" & strMyUserName & ".wota")
                lblGen(0).Caption = "Single Player"
                lblGen(1).Caption = "Multiplayer"
                lblGen(2).Caption = "Options"
                lblGen(3).Caption = "Credits"
                lblGen(4).Caption = "Exit"
                Exit Sub
            Else
                MsgBox "You must be logged in before loading your character!"
            End If
        End If
        
        If .Caption = "New Character" Then
            lblGen(0).Caption = "Single Player"
            lblGen(1).Caption = "Multiplayer"
            lblGen(2).Caption = "Options"
            lblGen(3).Caption = "Credits"
            lblGen(4).Caption = "Exit"
            bNewChar = True
            'Me.Hide
            frmSinglePlayer.Show
            Exit Sub
        End If
        If .Caption = "New User" Then
            lblGen(0).Caption = "Single Player"
            lblGen(1).Caption = "Multiplayer"
            lblGen(2).Caption = "Options"
            lblGen(3).Caption = "Credits"
            lblGen(4).Caption = "Exit"
            bNewChar = True
            'Me.Hide
            frmUser2.Show
            Exit Sub
        End If
        If .Caption = "Golden Sun Anonymous" Then
            frmBrowser.Show
            frmBrowser.Web.Navigate "http://www.doc-ent.com/gsa"
        End If
        If .Caption = "Game Options" Then
            lblGen(0).Caption = "Single Player"
            lblGen(1).Caption = "Multiplayer"
            lblGen(2).Caption = "Options"
            lblGen(3).Caption = "Credits"
            lblGen(4).Caption = "Exit"
            framOptions.Visible = True
        End If
    End With
End If
If Index = 2 Then
    With lblGen(2)
        If .Caption = "Options" Then
            lblGen(0).Caption = "Server Status"
            lblGen(1).Caption = "Game Options"
            lblGen(2).Caption = "Enter Code"
            lblGen(3).Caption = "View Ladder"
            lblGen(4).Caption = "Back"
            Exit Sub
        End If
        If .Caption = "Easter Egg Checklist" Then
            Call ShowEasterEgg
            framEasterEgg.Visible = True
            lblGen(0).Caption = "Single Player"
            lblGen(1).Caption = "Multiplayer"
            lblGen(2).Caption = "Options"
            lblGen(3).Caption = "Credits"
            lblGen(4).Caption = "Exit"
            Exit Sub
        End If
        If .Caption = "Enter Code" Then
            Dim strCode As String
            strCode = InputBox("Please enter your code", "Enter Code")
            'If strCode = "mazeload" Then
            '    MsgBox bMazeFirstLoad
            'End If
            'If strCode = "unloadbattle" Then
            '    BattleLoaded(1) = False
            '    BattleLoaded(2) = False
            'End If
            'If strCode = "battleloaded" Then
            '    MsgBox BattleLoaded(1)
            '    MsgBox BattleLoaded(2)
            'End If
            strCode = Eyncrypt(strCode)
            strCode = Mid$(strCode, 7, Len(strCode) - 12)
            Debug.Print vbNewLine
            Debug.Print strCode
            
            
            'If strCode = "–ﬁÊË" Then
            '    frmHost.Show
            'End If
            'If strCode = "join" Then
            '    frmJoin.Show
            'End If
            If strCode = "dear hex editor" Then
                MsgBox "Stop hex editing my code"
            End If
            'If strCode = "mute" Then
                'Launch Kill Mute window
            'End If
            If strCode = "rnhfd» ∆‰Ú‡Ëbd¬f⁄÷" Then '"97432decrypt12a3mk" Then
                'strload = InputBox("Enter character to load")
                'strField = InputBox("Enter field to decrypt")
                'strLength = InputBox("Enter length field to decrypt")
                'Call DecryptString(App.Path & "\" & strload & ".wota", strField, strLength)
            End If
            If Left$(strCode, 14) = "¬–rnhfd ‹∆‰Ú‡Ë" Then '"ah97432encrypt" Then
                Dim strEnc As String
                strEnc = Eyncrypt(Mid$(strCode, 15, Len(strCode)))
                strEnc = Mid$(strEnc, 7, Len(strEnc) - 12)
                Debug.Print strEnc
            End If
            
            If strCode = "ÿ ⁄Í‰“¬‹¬» ‡Ë" Then 'lemurianadept
                MsgBox "Piers unlocked!"
                frmUser2.cmbCharPic.AddItem "Piers"
            End If
            If strCode = "iamthepurplepeopleeater" Then
                MsgBox "Purple Piers unlocked!"
                frmUser2.cmbCharPic.AddItem "Purple Piers"
            End If
            If strCode = "∆∆∆Ó“‹‹ ‰"" Then 'cccwinner
                MsgBox "Caption Contest Character unlocked!"
                frmUser2.cmbCharPic.AddItem "Caption Contest Character"
            End If
            If strCode = "level" Then
                MsgBox strLvl
            End If
            If strCode = "⁄Ú∆ﬁﬁÿ» ¬»÷ ‹‹ÚÊﬁ∆“ ËÚ" Then
                MsgBox "Dead Kenny unlocked!"
                frmUser2.cmbCharPic.AddItem "Kenny"
            End If
            If strCode = "Ú ÿÿﬁÓË‰¬Ê–" Then
                MsgBox "All Isaac pictures changed to Yellow Isaac."
                frmMultiplayer.picChar(6).Picture = frmMultiplayer.picChar(15).Picture
                frmMultiplayer.picCharM(6).Picture = frmMultiplayer.picCharM(15).Picture
                frmMultiplayer.picChar(6).Width = frmMultiplayer.picChar(15).Width
                frmMultiplayer.picChar(6).Height = frmMultiplayer.picChar(15).Height
            End If
            If strCode = "thesewntogetherguy" Then
                MsgBox "Absolute K oS character unlocked!"
                frmUser2.cmbCharPic.AddItem "KOS"
            End If
            If strCode = "CurChar" Then
                MsgBox strChar(1)
                MsgBox strChar(2)
            End If
            If strCode = " Ë ÊËË–  ‹»Œ¬⁄  " Then
            '    frmEndGame.Show
            End If
            'If strCode = "wedontneednostinkinff7!" Then
            '    MsgBox "Cloud unlocked!"
            '    frmUser2.cmbCharPic.AddItem "Cloud"
            'End If
            'If strCode = "imlosingmymemory" Then
            '    MsgBox "Alternative log-in enabled."
            '    'cmdLogin.Enabled = True
            'End If
            If strCode = "Ë– ∆∆ »“Ëﬁ‰jb`" Then 'thecceditor510
                framCreate.Visible = True
            End If
            If strCode = "‹ Óƒ“ ¬ÿ ‰Ë" Then 'newbiealert
                MsgBox "Hello readers of the help file.  Congratulations on finding your first Easter Egg.", vbInformation, "Easter Egg #10"
                Call Encode("10", "EGG10", "EGGL10", App.Path & "\settings.ini")
            End If
            If strCode = "ÓﬁÿÃËÚ‡“‹Œ" Then 'wolftyping
                MsgBox "Wolf Avatars Unlocked!  Make sure that the select avatar screen is up right now or else this code will not work.  If it is not, just open the screen and re-enter the code.  Thanks for patronizing other Doc Entertainment software.", vbInformation, "Easter Egg #29"
                Call Encode("29", "EGG29", "EGGL29", App.Path & "\settings.ini")
                frmChat.imgSAvatar(14).Visible = True
                frmChat.imgSAvatar(15).Visible = True
                frmChat.imgSAvatar(16).Visible = True
                frmChat.imgSAvatar(17).Visible = True
            End If
            If strCode = "Ó“Ê ‰Ë–¬‹Ë–ﬁÍ"" Then 'wiserthanthou
                MsgBox "The Wise One unlocked!"
                frmUser2.cmbCharPic.AddItem "The Wise One"
            End If
            'If strCode = "morewinners" Then
            '    MsgBox "More scrambler winners"
            '    If frmUser2.User.State <> sckClosed Then
            '        frmChat.hSWinners.Max = 5
            '    End If
            'End If
            If strCode = "ÿ¬»» ‰ËﬁÍ‰‹" Then
                MsgBox "Version changed to ladder tourament."
                Version = "tournament"
            End If
            'If strCode = "Ë– Ê‡¬∆ ‹ ËﬁË– »ﬁﬁﬁ‰" Then
            '    MsgBox "Handicapp range extended!"
            '    frmHost.hHandicap.Min = -40
            '    frmHost.hHandicap.Max = 40
            '    frmJoin.hHandicap.Min = -40
            '    frmJoin.hHandicap.Max = 40
            '    frmHost.hHandicap.Visible = True
            '    frmJoin.hHandicap.Visible = True
            'End If
            If strCode = "rating" Then
                MsgBox strRating
            End If
            If strCode = "date" Then
                MsgBox strServerDate
            End If
            If Left$(strCode, 6) = "volume" Then
                Dim intVolume As Long
                intVolume = CLng(Mid$(strCode, 7, Len(strCode)))
                Midi.Volume = intVolume
            End If
            If Left$(strCode, 10) = "customchar" Then
                Dim intCustomCharPrint As Long
                intCustomCharPrint = CInt(Mid$(strCode, 11))
                With CustomChar(intCustomCharPrint)
                Debug.Print .Name
                Debug.Print .BaseAP
                Debug.Print .Picture
                End With
            End If
            If strCode = "MyStats" Then
                MsgBox "Characters - " & strChar(1) & " " & strChar(2) & vbNewLine & "Rating - " & strRating & vbNewLine & "Level - " & strLvl & vbNewLine & "Op. Rating - " & opRating & vbNewLine & "Op. Level - " & stroLvl & vbNewLine & "My Relative Level - " & RelativeLVL(1) & vbNewLine & "My Relative Rating - " & RelativeRating(1)
            End If
            If strCode = "GetLevel" Then
                intTemp = GetLevel(strRating)
                MsgBox intTemp
            End If
            If strCode = "GetDjinn" Then
                intTemp = GetDjinn(strRating)
                MsgBox intTemp
            End If
            If strCode = "Ë– Œﬁﬁ»Ã¬“‹ËÚﬁÍ‹Œ" Then
                On Error Resume Next
                lstEasterEgg.Clear
                framEasterEgg.Visible = True
                Dim strCheckEgg As String
                intEasterEggs = 0
                For i = 1 To 30 'Max easter eggs
                    strCheckEgg = UltraDecode("EGG" & CStr(i), "EGGL" & CStr(i), App.Path & "\settings.ini")
                    If strCheckEgg = CStr(i) Then
                        intEasterEggs = intEasterEggs + 1
                        lstEasterEgg.AddItem "Easter Egg " & CStr(i) & " - Found"
                    Else
                        lstEasterEgg.AddItem "Easter Egg " & CStr(i) & " - Not Found"
                    End If
                Next 'i
                If intEasterEggs >= 25 Then
                    MsgBox "Young Isaac and Young Garet unlocked!"
                    frmUser2.cmbCharPic.AddItem "Young Isaac"
                    frmUser2.cmbCharPic.AddItem "Young Garet"
                End If
            End If
        End If
    End With
End If
If Index = 3 Then
    With lblGen(3)
        If .Caption = "Credits" Then
            lblGen(0).Caption = "Roll Credits"
            lblGen(1).Caption = "Golden Sun Anonymous"
            lblGen(2).Caption = "Easter Egg Checklist"
            lblGen(3).Caption = ""
            lblGen(4).Caption = "Back"
            Exit Sub
        End If
        If .Caption = "View Ladder" Then
            MsgBox "The ladder isn't working/and I haven't tested it yet."
            'frmBrowser.Show
            'frmBrowser.Web.Navigate "http://" & IKILLKENNYIP & "/ladder.html"
        End If
    End With
End If
If Index = 4 Then
    With lblGen(4)
        If .Caption = "Exit" Then
            Unload Me
        End If
        If .Caption = "Back" Then
            lblGen(0).Caption = "Single Player"
            lblGen(1).Caption = "Multiplayer"
            lblGen(2).Caption = "Options"
            lblGen(3).Caption = "Credits"
            lblGen(4).Caption = "Exit"
            Exit Sub
        End If
    End With
End If
If Index = 5 Then
    frmBrowser.Show
    frmBrowser.Web.Navigate "http://www.google.com"
End If

End Sub
Sub SaveCharacter()
If strMyUserName = "" Then
    MsgBox "Can not save character!  You are not logged in to the server."
Else
    Dim strE As String
    Dim strLength As String
    Dim nsave As String
    
    nsave = App.Path & "\" & strMyUserName & ".wota"

    strE = Eyncrypt(strBS & strCoins & strBS2)
    strLength = Len(strCoins)
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "COINS1", "1", nsave)
    Call WriteIni("GEN", "COINS2", "0", nsave) 'Start with no coins
    
    strE = Eyncrypt(strBS & strLvl & strBS2)
    strLength = Len(strLvl)
    
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "LEVEL1", strLength, nsave)
    Call WriteIni("GEN", "LEVEL2", strE, nsave)

    strE = Eyncrypt(strBS & strRating & strBS2)
    strLength = Len(strRating)
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)

    Call WriteIni("GEN", "RATING1", strLength, nsave)
    Call WriteIni("GEN", "RATING2", strE, nsave)
        
    strE = Eyncrypt(strBS & strChar(1) & strBS2)
    strLength = Len(strChar(1))
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "CHAR1", strLength, nsave)
    Call WriteIni("GEN", "CHAR2", strE, nsave)
    
    strE = Eyncrypt(strBS & strItemName(intWeapon(1)) & strBS2)
    strLength = Len(strItemName(intWeapon(1)))
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "ITEMNAME1", strLength, nsave)
    Call WriteIni("GEN", "ITEMNAME2", strE, nsave)
    
    strE = Eyncrypt(strBS & strItemDamage(intWeapon(1)) & strBS2)
    strLength = Len(strItemDamage(intWeapon(1)))
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "ITEMDAMAGE1", strLength, nsave)
    Call WriteIni("GEN", "ITEMDAMAGE2", strE, nsave)

    strE = Eyncrypt(strBS & strItemDesc(intWeapon(1)) & strBS2)
    strLength = Len(strItemDesc(intWeapon(1)))
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "ITEMDESC1", strLength, nsave)
    Call WriteIni("GEN", "ITEMDESC2", strE, nsave)
    
    
    strE = Eyncrypt(strBS & strItemType(intWeapon(1)) & strBS2)
    strLength = Len(strItemType(intWeapon(1)))
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "ITEMTYPE1", strLength, nsave)
    Call WriteIni("GEN", "ITEMTYPE2", strE, nsave)
    
    
    strE = Eyncrypt(strBS & strItemSpcDesc(intWeapon(1)) & strBS2)
    strLength = Len(strItemSpcDesc(intWeapon(1)))
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "ITEMSNAME1", strLength, nsave)
    Call WriteIni("GEN", "ITEMSNAME2", strE, nsave)
    
    strE = Eyncrypt(strBS & strItemSpcDamage(intWeapon(1)) & strBS2)
    strLength = Len(strItemSpcDamage(intWeapon(1)))
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "ITEMSDAMAGE1", strLength, nsave)
    Call WriteIni("GEN", "ITEMSDAMAGE2", strE, nsave)
    
    strE = Eyncrypt(strBS & strItemSpcType(intWeapon(1)) & strBS2)
    strLength = Len(strItemSpcType(intWeapon(1)))
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "ITEMSTYPE1", strLength, nsave)
    Call WriteIni("GEN", "ITEMSTYPE2", strE, nsave)
    
    strE = Eyncrypt(strBS & strItemSpcDesc(intWeapon(1)) & strBS2)
    strLength = Len(strItemSpcDesc(intWeapon(1)))
    If Len(strLvl) < 10 Then
        strLength = "0" & strLength
    End If
    strLength = Eyncrypt(strLength)
    
    Call WriteIni("GEN", "ITEMSDESC1", strLength, nsave)
    Call WriteIni("GEN", "ITEMSDESC2", strE, nsave)
    
    For i = 1 To 30
        If strPsyName(i) <> "" Then
            Call Encode(strPsyName(i), "PSYNAME" & i, "PSYNAMEL" & i, nsave)
            Call Encode(strPsyDesc(i), "PSYDESC" & i, "PSYDESCL" & i, nsave)
            Call Encode(strPsyDamage(i), "PSYDAMAGE" & i, "PSYDAMAGEL" & i, nsave)
            Call Encode(strPsyType(i), "PSYTYPE" & i, "PSYTYPEL" & i, nsave)
            Call Encode(strPsyPP(i), "PSYPP" & i, "PSYPPL" & i, nsave)
            Call Encode(strPsyDjinn(i), "PSYDJINN" & i, "PSYDJINNL" & i, nsave)
        End If
        If i <= 10 Then
            If strDjinnName(i) <> "" Then
                Call Encode(strDjinnName(i), "DJINNNAME" & i, "DJINNNAMEL" & i, nsave)
                Call Encode(strDjinnType(i), "DJINNTYPE" & i, "DJINNTYPEL" & i, nsave)
                Call Encode(strDjinnDamage(i), "DJINNDAMAGE" & i, "DJINNDAMAGEL" & i, nsave)
                Call Encode(strDjinnDesc(i), "DJINNDESC" & i, "DJINNDESCL" & i, nsave)
            End If
            If strSumName(i) <> "" Then
                Call Encode(strSumName(i), "SUMNAME" & i, "SUMNAMEL" & i, nsave)
                Call Encode(strSumDesc(i), "SUMDESC" & i, "SUMDESCL" & i, nsave)
                Call Encode(strSumDjinn(i), "SUMDJINN" & i, "SUMDJINNL" & i, nsave)
            End If
        End If
    Next 'i
    
    If bCustomChar(1) <> 999 Then
        With CustomChar(bCustomChar(1))
            Call Encode(CStr(.BaseAP), "CAP", "CAPL", nsave)
            Call Encode(CStr(.BaseHP), "CHP", "CHPL", nsave)
            Call Encode(CStr(.BasePP), "CPP", "CPPL", nsave)
            Call Encode(CStr(.BaseDefense), "CDEFENSE", "CDEFENSEL", nsave)
            Call Encode(CStr(.BasePower), "CPOW", "CPOWL", nsave)
            Call Encode(CStr(.BaseRes), "CRES", "CRESL", nsave)
            Call Encode(CStr(.Name), "CNAME", "CNAMEL", nsave)
            Call Encode(CStr(.Picture), "CPIC", "CPICL", nsave)
            Call Encode("T", "CENABLED", "CENABLED", nsave)
        End With
    Else
        Call Encode("F", "CENABLED", "CENABLED", nsave)
    End If
End If

End Sub


Sub LoadCharacter(ByVal nsave As String)
On Error Resume Next
strNewCoins = UltraDecode("COINS1", "COINS2", nsave)
Dim iNewCoins As Long
iCoins = CLng(strCoins)
iNewCoins = CLng(strNewCoins)
iCoins = CLng(strCoins)
iCoins = iCoins + iNewCoins


If iNewCoins > 0 Then
    frmUser2.User.SendData "SINGLENAME" & strMyUserName & vbCrLf
    frmUser2.User.SendData "SINGLECOINS" & iNewCoins & vbCrLf
    Call Encode("0", "COINS1", "COINS2", nsave) 'Reset coins
Else
    MsgBox "Sorry, the character " & strMyUserName & " does currently not have any new coins to import.  Make sure that you saved the coins you earned in Single Player, that you have not gained any coins in the Multiplayer before importing the coins and make sure that " & strMyUserName & ".wota is in the same directory as this application.", vbCritical, "Error Loading Coins"
End If

End Sub


Private Sub lblVersion_Click()
If lblVersion.Caption = "Version 1.52 Beta 15 - 5/8/2004" Then lblVersion.Caption = "Version 1.33.7 Alpha 1 - 7/28/2010" Else lblVersion.Caption = "Version 1.52 Beta 15 - 5/8/2004"

End Sub

Private Sub lblVersion_DragDrop(Source As Control, X As Single, Y As Single)
MsgBox "You are probably asking yourself right now, what the hell did I just do?  You just found an easter egg (well, someone else probably found it and told you) that offers little to no function in the game.  However, here's a little secret: If you enter the password 'lemurianadept' in the cheats menu you may be pleasantly rewarded.", vbInformation, "Easter Egg #2"
Call Encode("2", "EGG2", "EGGL2", App.Path & "\settings.ini")
    
End Sub

Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    imgHuh.Visible = True
End If
End Sub


Private Sub Midi_EndOfStream(ByVal Result As Long)
Midi.Stop
Midi.Play
'Midi.Play
End Sub

Private Sub timeMidi_Timer()
On Error Resume Next
'Randomize numbers
Dim intRand As Long
Randomize
intRand = Int(Rnd * 1000)

strMusic = GetFromIni("GEN", "MUSIC", App.Path & "\settings.ini")
If strMusic = "ON" Then
    Dim temp As Integer
    Let temp = PeekMessage(curMessage, frmIntro.hWnd, 0, 0, PM_REMOVE)
    If Not (temp = 0 Or temp = -1) Then
        If curMessage.Message = MM_MCINOTIFY And (curMessage.wParam And MCI_NOTIFY_SUCCESSFUL) Then
            RepeatMidi
        End If
    End If
End If

End Sub
Sub ShowEasterEgg()
On Error Resume Next
lstEasterEgg.Clear
framEasterEgg.Visible = True
Dim strCheckEgg As String
intEasterEggs = 0
For i = 1 To 30 'Max easter eggs
    strCheckEgg = UltraDecode("EGG" & CStr(i), "EGGL" & CStr(i), App.Path & "\settings.ini")
    If strCheckEgg = CStr(i) Then
        intEasterEggs = intEasterEggs + 1
        lstEasterEgg.AddItem "Easter Egg " & CStr(i) & " - Found"
    Else
        lstEasterEgg.AddItem "Easter Egg " & CStr(i) & " - Not Found"
    End If
Next 'i
lblEggs.Caption = intEasterEggs
If intEasterEggs >= 10 And intEasterEggs < 20 Then
    lblEGen(0).Caption = "Congratulations!  You have found at least 10 easter eggs.  Here is your reward: Enter 'wiserthanthou' in the code window without quotes."
End If
If intEasterEggs >= 20 And intEasterEggs < 25 Then
    lblEGen(0).Caption = "You are a true master of the Easter Eggs (infact I've heard some call you the Easter Bunny... yeah, that was bad... I know :).  Here is a special reward."
    frmBrowser.Show
    frmBrowser.Web.Navigate "http://www.google.com"
End If
If intEasterEggs >= 25 And intEasterEggs < 30 Then
    lblEGen(0).Caption = "Congratulations, you've found 25 Easter Eggs.  For your reward, enter the password 'thegoodfaintyoung'.  Only 5 more to go!"
End If
If intEasterEggs = 30 Then
    lblEGen(0).Caption = "You have found all 30 Easter Eggs!  The final reward is meeting the creator..."
    frmBrowser.Show
    frmBrowser.Web.Navigate "http://www.google.com"
End If
End Sub
''Todo: encrypt crc and write it to a file.
 Public Function CRC32(ByRef bArrayIn() As Byte, ByVal lLen As Long) As Long
                Dim lCurPos As Long
                Dim lTemp As Long
                
                If lLen = 0 Then Exit Function 'In case of empty file
                  For lCurPos = 0 To lLen
                  lTemp = (((lTemp And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (crcTable((lTemp And 255) Xor bArrayIn(lCurPos)))
                Next lCurPos
                
                CRC32 = lTemp Xor &HFFFFFFFF
                'Returns CRC value
              End Function
              
              Public Function BuildTable() As Boolean
                Dim i As Long, X As Long, crc As Long
                Const Limit = &HEDB88320 'usally its shown backward, cant remember what it was.
                'Its the same polynomial that PKZIP uses (I Think)
                For i = 0 To 255
                  crc = i
                  For X = 0 To 7
                    If crc And 1 Then
                      crc = (((crc And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor Limit
                    Else
                      crc = ((crc And &HFFFFFFFE) \ 2) And &H7FFFFFFF
                    End If
                  Next X
                  crcTable(i) = crc
                Next i
              End Function
