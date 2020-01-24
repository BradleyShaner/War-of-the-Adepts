VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Configuration Tool"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1455
      Left            =   3960
      TabIndex        =   127
      Top             =   2640
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2566
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmServer.frx":0000
   End
   Begin VB.Frame framCustomChar 
      Caption         =   "Custom Character"
      Height          =   2175
      Left            =   2160
      TabIndex        =   96
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdCustSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   5280
         TabIndex        =   110
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   13
         Left            =   3720
         MaxLength       =   100
         TabIndex        =   126
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   12
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   124
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   11
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   122
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   10
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   121
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   9
         Left            =   3600
         MaxLength       =   150
         TabIndex        =   120
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   8
         Left            =   3600
         MaxLength       =   100
         TabIndex        =   119
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   7
         Left            =   3600
         MaxLength       =   1
         TabIndex        =   118
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   6
         Left            =   3600
         MaxLength       =   1
         TabIndex        =   117
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   5
         Left            =   480
         MaxLength       =   1
         TabIndex        =   116
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   4
         Left            =   480
         MaxLength       =   1
         TabIndex        =   115
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   3
         Left            =   480
         MaxLength       =   1
         TabIndex        =   114
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   2
         Left            =   360
         MaxLength       =   1
         TabIndex        =   113
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   1
         Left            =   840
         MaxLength       =   1
         TabIndex        =   112
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCustChar 
         Height          =   285
         Index           =   0
         Left            =   600
         MaxLength       =   20
         TabIndex        =   111
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblgen 
         AutoSize        =   -1  'True
         Caption         =   "Picture (no extension):"
         Height          =   195
         Index           =   52
         Left            =   3720
         TabIndex        =   125
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblgen 
         AutoSize        =   -1  'True
         Caption         =   "Luck:"
         Height          =   195
         Index           =   51
         Left            =   2520
         TabIndex        =   123
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label lblgen 
         AutoSize        =   -1  'True
         Caption         =   "Weakenss:"
         Height          =   195
         Index           =   50
         Left            =   1200
         TabIndex        =   109
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label lblgen 
         AutoSize        =   -1  'True
         Caption         =   "Strength:"
         Height          =   195
         Index           =   49
         Left            =   1200
         TabIndex        =   108
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lblgen 
         Caption         =   "Users Than Can Use Char (sep. by @):"
         Height          =   255
         Index           =   48
         Left            =   840
         TabIndex        =   107
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblgen 
         Caption         =   "Desc:"
         Height          =   255
         Index           =   47
         Left            =   3120
         TabIndex        =   106
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblgen 
         Caption         =   "Res:"
         Height          =   255
         Index           =   46
         Left            =   3120
         TabIndex        =   105
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblgen 
         Caption         =   "Pow:"
         Height          =   255
         Index           =   45
         Left            =   3120
         TabIndex        =   104
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblgen 
         Caption         =   "Def:"
         Height          =   255
         Index           =   44
         Left            =   120
         TabIndex        =   103
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblgen 
         Caption         =   "AP:"
         Height          =   255
         Index           =   43
         Left            =   120
         TabIndex        =   102
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblgen 
         Caption         =   "PP:"
         Height          =   255
         Index           =   42
         Left            =   120
         TabIndex        =   101
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblgen 
         Caption         =   "HP:"
         Height          =   255
         Index           =   41
         Left            =   120
         TabIndex        =   100
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblgen 
         Caption         =   "Class:"
         Height          =   255
         Index           =   40
         Left            =   120
         TabIndex        =   99
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblgen 
         Caption         =   "Name:"
         Height          =   255
         Index           =   39
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Timer timeFilePing 
      Interval        =   5000
      Left            =   5880
      Top             =   0
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   5
      Left            =   7440
      TabIndex        =   91
      Top             =   2280
      Width           =   615
   End
   Begin VB.Timer timeWait 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1800
      Top             =   6360
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   10
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   86
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   9
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   84
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   8
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   83
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   79
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Timer timeMulti 
      Interval        =   2500
      Left            =   4440
      Top             =   0
   End
   Begin VB.Timer timeServerPing 
      Interval        =   5000
      Left            =   5160
      Top             =   0
   End
   Begin VB.TextBox txtHTML 
      Height          =   2055
      Left            =   7800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   75
      Top             =   6000
      Width           =   1695
   End
   Begin VB.ListBox lstRank 
      Height          =   2010
      ItemData        =   "frmServer.frx":0098
      Left            =   7800
      List            =   "frmServer.frx":009A
      Sorted          =   -1  'True
      TabIndex        =   74
      Top             =   4440
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Chat 
      Index           =   0
      Left            =   720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9887
   End
   Begin VB.Timer timePing 
      Interval        =   5000
      Left            =   4800
      Top             =   0
   End
   Begin VB.ListBox lstStatus 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      ItemData        =   "frmServer.frx":009C
      Left            =   8520
      List            =   "frmServer.frx":009E
      TabIndex        =   65
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame framServerOptions 
      Caption         =   "Server Options"
      Height          =   3135
      Left            =   2280
      TabIndex        =   57
      Top             =   4080
      Width           =   5535
      Begin VB.CommandButton cmdAnnounce 
         Caption         =   "Announce"
         Height          =   255
         Left            =   4320
         TabIndex        =   130
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdateLadder 
         Caption         =   "Force Ladder Update"
         Height          =   495
         Left            =   4200
         TabIndex        =   129
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdShowCustChar 
         Caption         =   "Show/Hide Custom Char Editor"
         Height          =   495
         Left            =   2280
         TabIndex        =   97
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   255
         Left            =   2640
         TabIndex        =   95
         Top             =   2880
         Width           =   1095
      End
      Begin VB.FileListBox filChar 
         Height          =   480
         Left            =   3720
         Pattern         =   "*.gif;*"
         TabIndex        =   94
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkScrambler 
         Caption         =   "Scrambler Allowed"
         Height          =   255
         Left            =   1800
         TabIndex        =   89
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset Entire Ladder"
         Height          =   495
         Left            =   4200
         TabIndex        =   80
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdTOS 
         Caption         =   "Refresh TOS"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdCloseServer 
         Caption         =   "Go Down For Maitenece"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtVer 
         Height          =   285
         Left            =   840
         TabIndex        =   73
         Text            =   "1.33.7"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtContent 
         Height          =   285
         Left            =   1320
         TabIndex        =   71
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdMsg 
         Caption         =   "Msg"
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   255
         Left            =   3600
         TabIndex        =   68
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Left            =   3600
         TabIndex        =   67
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdMOTD 
         Caption         =   "Save"
         Height          =   255
         Left            =   2520
         TabIndex        =   64
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdKill 
         Caption         =   "Kill"
         Height          =   375
         Left            =   1200
         TabIndex        =   63
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtKill 
         Height          =   285
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtMOTD 
         Height          =   285
         Left            =   120
         MaxLength       =   350
         TabIndex        =   60
         Text            =   "Welcome to Golden Sun Anonymous' Online Battle Game Chat."
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmdGame 
         Caption         =   "Launch Game"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   58
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblIP 
         Height          =   255
         Left            =   3720
         TabIndex        =   88
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblgen 
         Caption         =   "0"
         Height          =   255
         Index           =   38
         Left            =   3480
         TabIndex        =   93
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblgen 
         Caption         =   "Sockets Open:"
         Height          =   255
         Index           =   37
         Left            =   3480
         TabIndex        =   92
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblgen 
         Caption         =   "IP:"
         Height          =   255
         Index           =   14
         Left            =   3720
         TabIndex        =   87
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblgen 
         Caption         =   "Version:"
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   72
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblKill 
         Caption         =   "Msg"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   70
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblgen 
         Caption         =   "Send Server Message:"
         Height          =   495
         Index           =   30
         Left            =   3600
         TabIndex        =   66
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblKill 
         Caption         =   "User"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   61
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblgen 
         Caption         =   "Set the Message of the Day:"
         Height          =   255
         Index           =   29
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox txtSummon 
      Height          =   285
      Index           =   3
      Left            =   600
      MaxLength       =   2
      TabIndex        =   56
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   54
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox txtSummon 
      Height          =   285
      Index           =   2
      Left            =   600
      MaxLength       =   1
      TabIndex        =   53
      Top             =   5160
      Width           =   255
   End
   Begin VB.TextBox txtSummon 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   51
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtSummon 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   49
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   46
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   45
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   44
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   4
      Left            =   7680
      TabIndex        =   42
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   3
      Left            =   7800
      TabIndex        =   40
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   2
      Left            =   7680
      TabIndex        =   38
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   1
      Left            =   7320
      MaxLength       =   1
      TabIndex        =   35
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   0
      Left            =   7320
      TabIndex        =   34
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   31
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   6
      Left            =   3000
      TabIndex        =   29
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   5
      Left            =   3360
      TabIndex        =   27
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   25
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   23
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   19
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   16
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   12
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   1
      Left            =   600
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Timer closeme 
      Interval        =   5000
      Left            =   5520
      Top             =   0
   End
   Begin VB.Timer timeUpdate 
      Interval        =   10000
      Left            =   7800
      Top             =   4080
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   0
      Left            =   1680
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9898
   End
   Begin MSWinsockLib.Winsock FileTransfer 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9880
   End
   Begin MSWinsockLib.Winsock nChat 
      Index           =   0
      Left            =   720
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9885
   End
   Begin VB.Label lblgen 
      Caption         =   "Server Version 1.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   53
      Left            =   0
      TabIndex        =   128
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label lblgen 
      Caption         =   "Add Mod:"
      Height          =   255
      Index           =   36
      Left            =   6720
      TabIndex        =   90
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblgen 
      Caption         =   "SPC % Occurence (Out of 100):"
      Height          =   435
      Index           =   35
      Left            =   3840
      TabIndex        =   85
      Top             =   1920
      Width           =   1365
   End
   Begin VB.Label lblgen 
      AutoSize        =   -1  'True
      Caption         =   "Multiply Mod:"
      Height          =   195
      Index           =   34
      Left            =   4560
      TabIndex        =   82
      Top             =   840
      Width           =   930
   End
   Begin VB.Label lblgen 
      AutoSize        =   -1  'True
      Caption         =   "Add Mod:"
      Height          =   195
      Index           =   33
      Left            =   4800
      TabIndex        =   81
      Top             =   480
      Width           =   690
   End
   Begin VB.Label lblgen 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Index           =   32
      Left            =   0
      TabIndex        =   78
      Top             =   3600
      Width           =   840
   End
   Begin VB.Label lblgen 
      Caption         =   "Djinn:"
      Height          =   255
      Index           =   28
      Left            =   0
      TabIndex        =   55
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Type:"
      Height          =   255
      Index           =   27
      Left            =   0
      TabIndex        =   52
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Description:"
      Height          =   255
      Index           =   26
      Left            =   0
      TabIndex        =   50
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Name:"
      Height          =   255
      Index           =   25
      Left            =   0
      TabIndex        =   48
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblgen 
      Caption         =   "Add Summon:"
      Height          =   255
      Index           =   24
      Left            =   0
      TabIndex        =   47
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label lbltypes 
      Caption         =   $"frmServer.frx":00A0
      Height          =   975
      Left            =   0
      TabIndex        =   43
      Top             =   7200
      Width           =   7695
   End
   Begin VB.Label lblgen 
      Caption         =   "Damage (%)"
      Height          =   255
      Index           =   23
      Left            =   6720
      TabIndex        =   41
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Attack Type:"
      Height          =   255
      Index           =   22
      Left            =   6720
      TabIndex        =   39
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Description:"
      Height          =   255
      Index           =   21
      Left            =   6720
      TabIndex        =   37
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Type:"
      Height          =   255
      Index           =   20
      Left            =   6720
      TabIndex        =   36
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Name:"
      Height          =   255
      Index           =   19
      Left            =   6720
      TabIndex        =   33
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Add Djinn:"
      Height          =   255
      Index           =   18
      Left            =   6720
      TabIndex        =   32
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblgen 
      Caption         =   "Damage:"
      Height          =   255
      Index           =   17
      Left            =   2400
      TabIndex        =   30
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblgen 
      Caption         =   "Type:"
      Height          =   255
      Index           =   16
      Left            =   2400
      TabIndex        =   28
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblgen 
      Caption         =   "SPC Base Damage:"
      Height          =   255
      Index           =   15
      Left            =   1920
      TabIndex        =   26
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblgen 
      Caption         =   "SPC Name:"
      Height          =   255
      Index           =   13
      Left            =   2400
      TabIndex        =   24
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Description:"
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   22
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Coins:"
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   21
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Name:"
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   18
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Add Item:"
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblgen 
      Caption         =   "Djinn:"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   15
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Rating:"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   14
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "PP:"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Damage:"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblgen 
      Caption         =   "Type:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Class:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Name:"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Add Psynergy:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblstatus 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblgen 
      Caption         =   "Status:"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCurUpdating As Boolean
Dim arrdatao As String
Dim upTime As Integer
Dim ladderRefresh As Integer
Dim SingleUser As Integer
Dim ladUser As Boolean
    Dim snewName As String
Dim strProfileUser As String
Dim strChangeUser As String
Dim strChangePass As String
Dim strNewPass As String
Dim strChangeRealPW As String

Dim iNewLogin As Long

Dim intLastPsy As Long
Dim intLastDjinn As Long
Dim intLastSummon As Long



Dim strNewUserPIN As String

Dim strSambarPath As String
Dim strNewPin As String

Dim intDjinnSaveHighScore As Long
Dim strDjinnSavePlayer As String

Dim bCreateCustChar As Boolean

Dim DontAdd As Boolean

Dim MaxCon As Integer
Dim UserName(0 To 21) As String
Dim UserRating(0 To 21) As String
Dim iNewServer As Integer
'Dim Game(1 To 20) As Games

Dim UserHP(1 To 20) As String

Dim ServerNum(1 To 20) As Long

Dim loginWait As Boolean

Dim strTOS(1 To 25) As String

Dim strTemp As String 'For file transfer
Dim strFileList As String

Dim ServerVersion As String
Dim UserVersion As String

Dim CurMsgName As String

Dim strMOTD As String

Dim etc As Integer

'Dim strdata As String
Dim IsLoaded(1 To 20) As Boolean
Dim noclose As Boolean
Dim NewUser As String
Dim NewPass As String
Dim IsItNew As String
Dim strTime As String
Dim strRating As String
Dim curTime As Long
Dim lstUser As String
Dim lstEmail As String
Dim realPassword As String
Dim strData() As String

Dim strChar(1 To 2) As String
Dim strCoins As String
Dim strWins As String
Dim strLoss As String
Dim strDisc As String
Dim strLvl As String
Dim strDjinn As String
Dim strType(1 To 2) As String
Dim strMyWeapon(1 To 2) As String

Dim strCurRating As String
Dim strCurLvl As String
Dim strCurCoins As String
Dim iCurRating As Integer
Dim iCurLvl As Integer
Dim iCurCoins As Integer
Dim iCurDjinn As Integer
Dim sCurDjinn As String
Dim curUser As String
Dim intUser As Integer
Dim curType As String

Dim strNewRank As String
Dim iNewRank As Integer
Dim iCurRank As Integer

Dim FileNameList() As String 'array of file names
Dim intCurFile As Integer 'what is the current
Dim b64 As New base64 'initiate base64 class
Dim GettingFile As Boolean
Dim File As String

Dim Users() As Users
Dim Game() As Game

Dim arrdata
Dim curPassword As String
'Dim curUser As String

Private Sub Chat_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
If Index = 0 Then
    'Beep
    iNewServer = 100
    For q = 1 To 20
    If Chat(q).State = sckClosed And iNewServer = 100 Then
        Chat(q).Accept requestID
        iNewServer = MaxCon
        Users(q).Enabled = True
        DoEvents
        Chat(q).SendData "ADMINTXT" & txtMOTD.Text & vbCrLf
        Chat(q).SendData "ISAACNUM" & q & vbCrLf
        If chkScrambler.Value = 1 Then
            Chat(q).SendData "SCRAMON" & vbCrLf
        Else
            Chat(q).SendData "SCRAMOFF" & vbCrLf
        End If
            
        Exit Sub
    End If
    Next 'i
End If 'if index = 0
End Sub

Private Sub Chat_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo err
    Dim nfile As String
    Dim xfile As String
    'Change for the ladder tournament:
    nfile = App.Path & "\user.ini"
    xfile = App.Path & "\data.ini"
Dim strtime2 As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
strtime2 = Format(Now, "dd-mmmm")
'strTime = CStr(curTime)
Dim strdatao As String
Chat(Index).GetData strdatao
strData = Split(strdatao, vbCrLf, -1, vbTextCompare)
strData = strData

DontAdd = False

For i = 0 To UBound(strData)
    'Changed for the Ladder Tournament
    nSave = App.Path & "\user.ini"
    'If i > UBound(strData) Then Exit Sub
    If Left$(strData(i), 7) = "CHATTXT" Then
    Dim strChatTxt As String
    strChatTxt = Mid(strData(i), 8, Len(strData(i)))
    txtChat.Text = txtChat.Text & vbNewLine & strChatTxt
    Call WriteIni("CHAT", strTime & Chat(Index).RemoteHostIP, strChatTxt, App.Path & "\" & strtime2 & ".ini")
    Call SendChat(strChatTxt)
    End If
    If Left$(strData(i), 7) = "GOLDTXT" Then
    strChatTxt = Mid(strData(i), 8, Len(strData(i)))
    txtChat.Text = txtChat.Text & vbNewLine & strChatTxt
    Call WriteIni("GOLD", strTime & Chat(Index).RemoteHostIP, strChatTxt, App.Path & "\" & strtime2 & ".ini")
    For q = 1 To 20
        If Users(q).Enabled = True Then
            Chat(q).SendData "GOLDTXT" & strChatTxt & vbCrLf
            DoEvents
        End If
    Next 'i
    End If
    If Left$(strData(i), 6) = "LOGOFF" Then
        Users(Index).Enabled = False
        Chat(Index).Close
        Game(Index).Enabled = False
        Call KillChar(CStr(Index))
        Exit Sub
    End If
    If Left$(strData(i), 16) = "LADDERTOURNAMENT" Then
        Dim strLadT As String
        strLadT = Mid(strData(i), 17, Len(strData(i)))
        'Call WriteIni("TOURNAMENT", strTime & Chat(Index).RemoteHostIP & Users(Index).Name, strLadT, App.Path & "\" & strtime2 & ".ini")
        If Mid$(strData(i), 17, 4) = "IWIN" Then
            Call AdminChat(Users(Index).Name & " lost!")
        End If
        If Mid$(strData(i), 17, 5) = "ILOST" Then
            Call AdminChat(Users(Index).Name & " won!")
        End If
        Dim strSendLadder
        strSendLadder = strLadT
        strtemps = Split(strLadT, ":", -1, vbTextCompare)
        If Left$(CStr(strtemps(1)), 5) = " MYHP" Then
            UserHP(Index) = Mid$(strtemps(1), 6, Len(strtemps(1)))
        End If
        If strtemps(1) <> " BATTLELOADED" Then
            'Call SendChat(strSendLadder)
        Else
            Debug.Print strtemps(1)
        End If
    End If
    If Left$(strData(i), 5) = "GETHP" Then
        Dim strGetHP As String
        strGetHP = Mid$(strData(i), 6, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = strGetHP Then
                Call AdminChat(strGetHP & "'s current HP is: " & UserHP(q))
            End If
        Next 'q
    End If
    If Left$(strData(i), 5) = "GETPP" Then
        Dim strGetPP As String
        strGetPP = Mid$(strData(i), 6, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = strGetPP Then
                Chat(q).SendData "GETPP"
            End If
        Next 'q
    End If
    If Left$(strData(i), 5) = "GETAP" Then
        Dim strGetAP As String
        strGetAP = Mid$(strData(i), 6, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = strGetAP Then
                Chat(q).SendData "GETAP"
            End If
        Next 'q
    End If
    If Left$(strData(i), 7) = "MODWARN" Then
        Dim strKarmaUser As String
        strKarmaUser = Mid$(strData(i), 8, Len(strData(i)))
        
    
        Dim strKarma As String
        strKarma = GetFromIni(strKarmaUser, "KARMA", nfile)
        If strKarma = "" Then strKarma = "0"
        strKarma = CStr(CInt(strKarma) - 1)
        Call WriteIni(strKarmaUser, "KARMA", strKarma, nfile)
        For q = 1 To 20
            If Users(q).Name = strKarmaUser Then
                Chat(q).SendData "MODWARN"
            End If
        Next 'i
    End If
    If Left$(strData(i), 9) = "MODPRAISE" Then
        'Dim strKarmaUser As String
        strKarmaUser = Mid$(strData(i), 10, Len(strData(i)))
        
    
        'Dim strKarma As String
        strKarma = GetFromIni(strKarmaUser, "KARMA", nfile)
        If strKarma = "" Then strKarma = "0"
        strKarma = CStr(CInt(strKarma) + 1)
        Call WriteIni(strKarmaUser, "KARMA", strKarma, nfile)
        For q = 1 To 20
            If Users(q).Name = strKarmaUser Then
                Chat(q).SendData "MODPRAISE"
            End If
        Next 'i
    End If
        
    If Left$(strData(i), 7) = "GETTIME" Then
        Chat(Index).SendData "CHATMSG" & "The current server time is: " & strTime & vbCrLf
    End If
    If Left$(strData(i), 6) = "SHOWIP" Then
        Dim strIPUser As String
        strIPUser = Mid$(strData(i), 7, Len(strData(i)))
        strIPUser = strIPUser & "'s IP is " & Chat(Index).RemoteHostIP
        Call SendMeChat(strIPUser)
    End If
    If Left$(strData(i), 9) = "GETJOINIP" Then
        Dim strJoinIPUser As String
        strJoinIPUser = Mid$(strData(i), 10, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Enabled = True And Users(q).Name = strJoinIPUser Then
                Chat(Index).SendData "GETJOINIP" & Chat(q).RemoteHostIP & vbCrLf
                Exit For
            End If
        Next 'q
    End If
    If Left$(strData(i), 9) = "LADDERWIN" Then
        Call AdminChat(Users(Index).Name & " won!")
    End If
    If Left$(strData(i), 10) = "LADDERLOSS" Then
        Call AdminChat(Users(Index).Name & " lost!")
    End If
    If Left$(strData(i), 7) = "IAMAWAY" Then
        Call SendDataToAll("AVATAR" & CStr(Index) & "@" & "Away")
        Call SendDataToAll("UPDATEDISPLAY")
        Users(Index).Avatar = "Away"
        Users(Index).Away = True
    End If
    If Left$(strData(i), 6) = "AVATAR" Then
        Users(Index).Avatar = Mid$(strData(i), 7, Len(strData(i)))
        Call SendDataToAll("AVATAR" & CStr(Index) & "@" & Users(Index).Avatar)
        Call SendDataToAll("UPDATEDISPLAY")
    End If
    If Left$(strData(i), 7) = "NOTAWAY" Then
        Dim strAwayUser As String
        strAwayUser = Mid$(strData(i), 8, Len(strData(i)))
        strAwayUser = GetFromIni(strAwayUser, "AVATAR", App.Path & "\user.ini")
        Users(Index).Avatar = strAwayUser
        Users(Index).Away = False
        Call SendDataToAll("AVATAR" & CStr(Index) & "@" & strAwayUser)
        Call SendDataToAll("UPDATEDISPLAY")
    End If
    If Left$(strData(i), 5) = "MODON" Then
        Call SendDataToAll(strData(i))
        Call SendDataToAll("UPDATEDISPLAY")
    End If
    If Left$(strData(i), 7) = "ADMINON" Then
        Call SendDataToAll(strData(i))
        Call SendDataToAll("UPDATEDISPLAY")
    End If
    
    
    
    If Left$(strData(i), 5) = "METXT" Then
'        Dim strChatTxt As String
        strChatTxt = Mid(strData(i), 6, Len(strData(i)))
        'If strChatTxt = "dragoon has hosted a game (192.168.0.2)" Then strChatTxt = "dragoon has hosted a game (68.60.228.15)"
        txtChat.Text = txtChat.Text & vbNewLine & strChatTxt
        Call WriteIni("ME", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
        Call SendMeChat(strChatTxt)
    End If
    If Left$(strData(i), 10) = "CHATFREEZE" Then
'        Dim strChatTxt As String
        Dim strFreezeName As String
        strFreezeName = Mid(strData(i), 11, Len(strData(i)))
        Call WriteIni("FREEZE", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
        For q = 1 To 20
            If Users(q).Name = strFreezeName Then
                Chat(q).SendData "CHATFREEZE" & vbCrLf
                DoEvents
            End If
            DoEvents
        Next 'q
    End If
    If Left$(strData(i), 8) = "60FREEZE" Then
'        Dim strChatTxt As String
'        Dim strFreezeName As String
        strFreezeName = Mid(strData(i), 9, Len(strData(i)))
        Call WriteIni("FREEZE", strFreezeName, strChatTxt, App.Path & "\" & strtime2 & ".ini")
        For q = 1 To 20
            If Users(q).Name = strFreezeName Then
                Chat(q).SendData "60FREEZE" & vbCrLf
                DoEvents
            End If
            DoEvents
        Next 'q
    End If
    
    If Left$(strData(i), 5) = "PUSER" Then
        strProfileUser = Mid$(strData(i), 6, Len(strData(i)))
    End If
    
    Dim pSave As String
    pSave = App.Path & "\user.ini"
    
    If Left$(strData(i), 5) = "PNAME" Then
        Dim strTempProfile As String
        strTempProfile = Mid$(strData(i), 6, Len(strData(i)))
        Call WriteIni(strProfileUser, "PNAME", strTempProfile, pSave)
    End If
    If Left$(strData(i), 4) = "PAGE" Then
        strTempProfile = Mid$(strData(i), 5, Len(strData(i)))
        Call WriteIni(strProfileUser, "PAGE", strTempProfile, pSave)
    End If
    If Left$(strData(i), 4) = "PSEX" Then
        strTempProfile = Mid$(strData(i), 5, Len(strData(i)))
        Call WriteIni(strProfileUser, "PSEX", strTempProfile, pSave)
    End If
    If Left$(strData(i), 9) = "PLOCATION" Then
        strTempProfile = Mid$(strData(i), 10, Len(strData(i)))
        Call WriteIni(strProfileUser, "PLOCATION", strTempProfile, pSave)
    End If
    If Left$(strData(i), 4) = "PAIM" Then
        strTempProfile = Mid$(strData(i), 5, Len(strData(i)))
        Call WriteIni(strProfileUser, "PAIM", strTempProfile, pSave)
    End If
    If Left$(strData(i), 4) = "PMSN" Then
        strTempProfile = Mid$(strData(i), 5, Len(strData(i)))
        Call WriteIni(strProfileUser, "PMSN", strTempProfile, pSave)
    End If
    If Left$(strData(i), 4) = "PICQ" Then
        strTempProfile = Mid$(strData(i), 5, Len(strData(i)))
        Call WriteIni(strProfileUser, "PICQ", strTempProfile, pSave)
    End If
    If Left$(strData(i), 6) = "PEMAIL" Then
        strTempProfile = Mid$(strData(i), 7, Len(strData(i)))
        Call WriteIni(strProfileUser, "PEMAIL", strTempProfile, pSave)
    End If
    If Left$(strData(i), 6) = "POTHER" Then
        strTempProfile = Mid$(strData(i), 7, Len(strData(i)))
        Call WriteIni(strProfileUser, "POTHER", strTempProfile, pSave)
    End If
    If Left$(strData(i), 4) = "PURL" Then
        strTempProfile = Mid$(strData(i), 5, Len(strData(i)))
        Call WriteIni(strProfileUser, "PURL", strTempProfile, pSave)
    End If
    If Left$(strData(i), 7) = "PAVATAR" Then
        strTempProfile = Mid$(strData(i), 8, Len(strData(i)))
        Call WriteIni(strProfileUser, "AVATAR", strTempProfile, pSave)
        Dim intUpdateAv As Long
        For q = 1 To 20
            If Users(q).Enabled = True And Users(q).Name = strProfileUser Then
                intUpdateAv = q
                Exit For
            End If
        Next 'q
        Users(intUpdateAv).Avatar = strTempProfile
        Call SendDataToAll("AVATAR" & CStr(intUpdateAv) & "@" & strTempProfile)
        Call SendDataToAll("UPDATEDISPLAY")
    End If
    
    If Left$(strData(i), 10) = "GETPROFILE" Then
        strProfileUser = Mid$(strData(i), 11, Len(strData(i)))
        Chat(Index).SendData "PUSER" & strProfileUser & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "PNAME", pSave)
        Chat(Index).SendData "PNAME" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "PAGE", pSave)
        Chat(Index).SendData "PAGE" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "PSEX", pSave)
        Chat(Index).SendData "PSEX" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "PLOCATION", pSave)
        Chat(Index).SendData "PLOCATION" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "PAIM", pSave)
        Chat(Index).SendData "PAIM" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "PMSN", pSave)
        Chat(Index).SendData "PMSN" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "PEMAIL", pSave)
        Chat(Index).SendData "PEMAIL" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "POTHER", pSave)
        Chat(Index).SendData "POTHER" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "WINS", pSave)
        Chat(Index).SendData "PWINS" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "LOSS", pSave)
        Chat(Index).SendData "PLOSS" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "RATING", pSave)
        Chat(Index).SendData "PRATING" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "COINS", pSave)
        Chat(Index).SendData "PCOINS" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "CHAR", pSave)
        Chat(Index).SendData "PCHARACTER" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "PICQ", pSave)
        Chat(Index).SendData "PICQ" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "KARMA", pSave)
        Chat(Index).SendData "PKARMA" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "AVATAR", pSave)
        Chat(Index).SendData "PAVATAR" & strTempProfile & vbCrLf
        strTempProfile = GetFromIni(strProfileUser, "PURL", pSave)
        Chat(Index).SendData "PURL" & strTempProfile & vbCrLf
        
    End If
'    If Left$(strdata(i), 7) = "CHATSND" Then
'        Dim strChatTxt As String
'        strChatTxt = Mid(strdata(i), 8, Len(strdata(i)))
'        Call WriteIni("SOUND", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
'        Call SendChat(Users(Index).Name & " played sound: " & strChatTxt)
'        Call SendSndChat(strChatTxt)
'    End If
    
    If Left$(strData(i), 8) = "CHATNAME" Then
        Users(Index).Name = Mid(strData(i), 9, Len(strData(i)))
        Users(Index).IP = Server(Index).RemoteHostIP
        Users(Index).Avatar = GetFromIni(Users(Index).Name, "AVATAR", App.Path & "\user.ini")
        If Users(Index).Avatar = "" Then
            Call WriteIni(Users(Index).Name, "AVATAR", "Ivan", App.Path & "\user.ini")
            Users(Index).Avatar = "Ivan"
        End If
        
        Dim strCurUser As String
        strCurUser = Users(Index).Name
        Call WriteIni(strCurUser, "IP", Users(Index).IP, App.Path & "\user.ini")
        Chat(Index).SendData "REALIP" & Chat(Index).RemoteHostIP & vbCrLf
        Call SendDataToAll("USERNAME" & CStr(Index) & "@" & Users(Index).Name)
        Call SendDataToAll("AVATAR" & CStr(Index) & "@" & Users(Index).Avatar)
    End If
    
    If Left$(strData(i), 9) = "WHERERQST" Then
        Dim arrtmp As String
        arrtmp = Mid$(strData(i), 10, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = arrtmp And arrtmp <> "" Then
                Chat(Index).SendData "WHERE" & CStr(q) & vbCrLf
            End If
        Next
    End If
    If Left$(strData(i), 6) = "NEEDIP" Then
        Dim findIP As String
        findIP = Mid$(strData(i), 7, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = findIP And findIP <> "" Then
                Chat(Index).SendData "MODTXT" & Users(q).Name & "'s IP is " & Users(q).IP & vbCrLf
            End If
        Next
    End If
    
'    If Left$(strdata(i), 9) = "WHERERPLY" Then
'        Dim tmp2 As String
'        arrtmp = Split(Mid$(strdata(i), 10, Len(strdata(i))), " ", -1, vbTextCompare)
'        For q = 1 To UBound(arrtmp)
'            tmp2 = tmp2 + arrtmp(q)
'        Next q
'        For q = 1 To 20
'            If Users(q).Name = arrtmp(0) Then
'                Chat(q).SendData "WHEREBACK" & tmp2 & vbCrLf
'            End If
'        Next
'    End If
    
    
    If Left$(strData(i), 11) = "CHATMSGNAME" Then
        CurMsgName = Mid(strData(i), 12, Len(strData(i)))
    End If
    
    If Left$(strData(i), 11) = "CHATMSGTEXT" Then
        Dim CurMsgText As String
        CurMsgText = Users(Index).Name & ": " & Mid(strData(i), 12, Len(strData(i)))
        Call ChatMsg(CurMsgText, CurMsgName)
        Call WriteIni("MESSAGES", strTime, CurMsgName & ": " & CurMsgText & "  to " & Users(Index).Name, App.Path & "\" & strtime2 & ".ini")
    End If
    
    If Left$(strData(i), 8) = "ADMINTXT" Then
        Dim strAdminChatTxt As String
        strChatTxt = Mid(strData(i), 9, Len(strData(i)))
        txtChat.Text = txtChat.Text & vbNewLine & strChatTxt
        Call WriteIni("ADMINCHAT", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
        Call AdminChat(strChatTxt)
    End If
    If Left$(strData(i), 6) = "MODTXT" Then
        strChatTxt = Mid(strData(i), 7, Len(strData(i)))
        txtChat.Text = txtChat.Text & vbNewLine & strChatTxt
        Call WriteIni("MODCHAT", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
        Call ModChat(strChatTxt)
    End If
    If Left$(strData(i), 7) = "TALKTXT" Then
        strChatTxt = Mid(strData(i), 8, Len(strData(i)))
        txtChat.Text = txtChat.Text & vbNewLine & strChatTxt
        Call WriteIni("CHAT", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
        Call TalkChat(strChatTxt)
    End If
    If Left$(strData(i), 7) = "GAMETXT" Then
        strChatTxt = Mid(strData(i), 8, Len(strData(i)))
        txtChat.Text = txtChat.Text & vbNewLine & strChatTxt
        Call WriteIni("MODCHAT", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
        Call GameChat(strChatTxt, False)
    End If
    If Left$(strData(i), 9) = "GAMEMETXT" Then
        strChatTxt = Mid(strData(i), 10, Len(strData(i)))
        txtChat.Text = txtChat.Text & vbNewLine & strChatTxt
        Call WriteIni("MODCHAT", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
        Call GameChat(strChatTxt, True)
    End If
    
    
    
    
    If Left$(strData(i), 9) = "GETRATING" Then
        Dim strCheckRating As String
        strCheckRating = Mid(strData(i), 10, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Name = strCheckRating Then
                Chat(Index).SendData "CHATRATING" & Users(q).Rating & vbCrLf
            End If
        Next 'q
    End If

    If Left$(strData(i), 5) = "GETIP" Then
        Dim strCheckIP As String
        strCheckIP = Mid(strData(i), 6, Len(strData(i)))
        Dim strTempIP As String
        strTempIP = GetFromIni(strCheckIP, "IP", App.Path & "\user.ini")
        Chat(Index).SendData "CHATTXT" & strCheckIP & "'s IP is " & strTempIP & vbCrLf

    End If
    
    'If Left$(strData(i), 7) = "IAMAWAY" Then
    '    Users(Index).Away = True
    'End If
    'If Left$(strData(i), 7) = "NOTAWAY" Then
    '    Users(Index).Away = False
    'End If
    
    If Left$(strData(i), 8) = "ISAACPIC" Then
        Users(Index).Pic = Mid(strData(i), 9, Len(strData(i)))
        Users(Index).Left = 0
        Users(Index).Top = 0
        Call AddMUser(Index, True)
    End If
    If Left$(strData(i), 6) = "SCREEN" Then
        Users(Index).Screen = Mid(strData(i), 7, Len(strData(i)))
        DontAdd = True
    End If
    If Left$(strData(i), 6) = "ISAACX" Then
        Users(Index).Left = Mid(strData(i), 7, Len(strData(i)))
        DontAdd = True
    End If
    If Left$(strData(i), 6) = "ISAACY" Then
        Users(Index).Top = Mid(strData(i), 7, Len(strData(i)))
        DontAdd = True
    End If
    
    If Left$(strData(i), 10) = "CREATEGAME" Then
        Game(Index).Name = Mid(strData(i), 11, Len(strData(i)))
        Game(Index).Enabled = True
        Game(Index).IP = Chat(Index).RemoteHostIP
    End If
    If Left$(strData(i), 8) = "GAMEHOST" Then
        Game(Index).Host = Mid(strData(i), 9, Len(strData(i)))
    End If
    If Left$(strData(i), 11) = "GETGAMELIST" Then
        Call GetGameList(Index)
    End If
    If Left$(strData(i), 9) = "CLOSEGAME" Then
        Game(Index).Enabled = False
    End If
    If Left$(strData(i), 8) = "JOINGAME" Then
        strgame = Mid(strData(i), 9, Len(strData(i)))
        Call JoinGame(strgame, Index)
    End If
    
    If Left$(strData(i), 10) = "SWITCHCHAR" Then
        
        Dim strnewchar As String
        strnewchar = Mid(strData(i), 11, Len(strData(i)))
        strChar(1) = strnewchar
        Dim sTotal As String
        sTotal = Users(Index).Name
        
        If strChar(1) = "Isaac" Or strChar(1) = "Guard" Or strChar(1) = "Gladiator" Then
         Call WriteIni(sTotal, "TYPE", "E", nfile)
        End If
        If strChar(1) = "Kenny" Or strChar(1) = "Jenna" Or strChar(1) = "Garret" Or strChar(1) = "Saturos" Or strChar(1) = "Menardi" Then
            Call WriteIni(sTotal, "TYPE", "F", nfile)
        End If
        If strChar(1) = "Ivan" Or strChar(1) = "Sheba" Or strChar(1) = "Cloud" Then
            Call WriteIni(sTotal, "TYPE", "N", nfile)
        End If
        If strChar(1) = "Purple Piers" Or strChar(1) = "Piers" Or strChar(1) = "Mia" Or strChar(1) = "Alex" Or strChar(1) = "Caption Contest Character" Then
            Call WriteIni(sTotal, "TYPE", "W", nfile)
        End If
        If strChar(1) = "Felix" Or strChar(1) = "The Wise One" Then
            Call WriteIni(sTotal, "TYPE", "H", nfile)
        End If
        If strChar(1) = "Kraden" Or strChar(1) = "KOS" Or strChar(1) = "Karst" Or strChar(1) = "Agiato" Then
            Call WriteIni(sTotal, "TYPE", "D", nfile)
        End If
    
        Call WriteIni(sTotal, "CHAR", strnewchar, nfile)
    
    End If
    
    If Left$(strData(i), 11) = "2SWITCHCHAR" Then
        

        strnewchar = Mid(strData(i), 12, Len(strData(i)))
        strChar(1) = strnewchar
        sTotal = Users(Index).Name
        
        If strChar(1) = "Isaac" Or strChar(1) = "Guard" Or strChar(1) = "Gladiator" Then
         Call WriteIni(sTotal, "TYPE2", "E", nfile)
        End If
        If strChar(1) = "Kenny" Or strChar(1) = "Jenna" Or strChar(1) = "Garret" Or strChar(1) = "Saturos" Or strChar(1) = "Menardi" Then
            Call WriteIni(sTotal, "TYPE2", "F", nfile)
        End If
        If strChar(1) = "Ivan" Or strChar(1) = "Sheba" Or strChar(1) = "Cloud" Then
            Call WriteIni(sTotal, "TYPE2", "N", nfile)
        End If
        If strChar(1) = "Purple Piers" Or strChar(1) = "Piers" Or strChar(1) = "Mia" Or strChar(1) = "Alex" Or strChar(1) = "Caption Contest Character" Then
            Call WriteIni(sTotal, "TYPE2", "W", nfile)
        End If
        If strChar(1) = "Felix" Or strChar(1) = "The Wise One" Then
            Call WriteIni(sTotal, "TYPE2", "H", nfile)
        End If
        If strChar(1) = "Kraden" Or strChar(1) = "KOS" Or strChar(1) = "Karst" Or strChar(1) = "Agiato" Then
            Call WriteIni(sTotal, "TYPE2", "D", nfile)
        End If
    
        Call WriteIni(sTotal, "CHAR2", strnewchar, nfile)
    
    End If
    
    If Left$(strData(i), 14) = "SWITCHCUSTCHAR" Then
        'Dim strChar As String
        strChar(1) = Mid(strData(i), 15, Len(strData(i)))
        strChar(1) = strChar(1)
        'Dim sTotal As String
        sTotal = Users(Index).Name
        
        Dim strCustCharType As String
        strCustCharType = GetFromIni(strChar(1), "CLASS", App.Path & "\customchar.ini")
        Call WriteIni(sTotal, "TYPE", strCustCharType, nfile)
        strCustCharType = GetFromIni(strChar(1), "NAME", App.Path & "\customchar.ini")
        Call WriteIni(sTotal, "CHAR", strCustCharType, nfile)
        
    End If
    
    If Left$(strData(i), 9) = "KILLCOINS" Then
        
        Dim strNewCoins As String
        Dim strCurKillCoins As String
        strNewCoins = Mid(strData(i), 10, Len(strData(i)))
        strCurKillCoins = GetFromIni(Users(Index).Name, "COINS", nfile)
        strNewCoins = CInt(strCurKillCoins - strNewCoins)
        Call WriteIni(Users(Index).Name, "COINS", strNewCoins, nfile)
    
    End If
    
    If Left$(strData(i), 10) = "INGAMECHAT" Then
        Dim strInChat As String
        strInChat = Mid(strData(i), 11, Len(strData(i)))
        Call InGameChat(Index, strInChat)
    End If
    
    If Left$(strData(i), 7) = "COMPBAN" Then
        strBan = Mid(strData(i), 8, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Enabled = True And Users(q).Name = strBan Then
                Dim bSave As String
                Dim iBanMax As Integer
                bSave = App.Path & "\ban.ini"
                iBanMax = GetFromIni("GEN", "MAX", bSave)
                Call WriteIni("GEN", CStr(iBanMax + 1), Users(q).IP, bSave)
                Call WriteIni("GEN", "MAX", CStr(iBanMax + 1), bSave)
                Chat(q).SendData "CHATBAN" & vbCrLf
            End If
        Next 'q
    End If
    If Left$(strData(i), 5) = "IPBAN" Then
 

        strBan = Mid(strData(i), 6, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Enabled = True And Users(q).Name = strBan Then
                bSave = App.Path & "\ban.ini"
                iBanMax = GetFromIni("GEN", "MAX", bSave)
                Call WriteIni("GEN", CStr(iBanMax + 1), Chat(q).RemoteHostIP, bSave)
                Call WriteIni("GEN", "MAX", CStr(iBanMax + 1), bSave)
                Chat(q).SendData "KILL" & vbCrLf
            End If
        Next 'q
    End If
    
    If Left$(strData(i), 6) = "GETPIN" Then
        Dim strGetPin As String
        Dim strPinName As String
        strPinName = Mid(strData(i), 7, Len(strData(i)))
        strGetPin = GetFromIni(strPinName, "PINNUM", App.Path & "\user.ini")
        Chat(Index).SendData "CHATMSG" & strPinName & "'s PIN Number is " & strGetPin & vbCrLf
    End If
    If Left$(strData(i), 6) = "PINBAN" Then
        strPinName = Mid(strData(i), 7, Len(strData(i)))
        strGetPin = GetFromIni(strPinName, "PINNUM", App.Path & "\user.ini")
        For q = 1 To 20
            If Users(q).Enabled = True And Users(q).Name = strPinName Then
                Chat(q).SendData "PINBAN" & vbCrLf
            End If
        Next 'q
        Dim intPinBanMax As Integer
        intPinBanMax = CInt(GetFromIni("GEN", "TOTAL", App.Path & "\pinban.ini"))
        intPinBanMax = intPinBanMax + 1
        Call WriteIni("GEN", "TOTAL", CStr(intPinBanMax), App.Path & "\pinban.ini")
        Call WriteIni("GEN", CStr(intPinBanMax), strGetPin, App.Path & "\pinban.ini")
    End If
    
    If Left$(strData(i), 9) = "ADMINKILL" Then
        Dim strKill As String
        strKill = Mid(strData(i), 10, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Enabled = True And Users(q).Name = strKill Then
                Chat(q).SendData "KILL" & vbCrLf
            End If
        Next 'q
        txtChat.Text = txtChat.Text & vbNewLine & strKill & " was kicked by " & Users(Index).Name
    End If
    
    If Left$(strData(i), 9) = "ADMINKICK" Then
        Dim strKick As String
        strKick = Mid(strData(i), 10, Len(strData(i)))
        For q = 1 To 20
            If Users(q).Enabled = True And Users(q).Name = strKick Then
                Chat(q).SendData "KICK" & vbCrLf
            End If
        Next 'q
        txtChat.Text = txtChat.Text & vbNewLine & strKill & " was kicked by " & Users(Index).Name
    End If
    If Left$(strData(i), 10) = "CHATREPORT" Then
        Dim strReport As String
        strReport = Mid(strData(i), 11, Len(strData(i)))
        Dim strReport2
        strReport2 = Split(strReport, "@", -1, vbTextCompare)
        Call WriteIni("GEN", strTime, "Narc on user " & strReport2(1) & " by user " & Users(Index).Name & " for " & CStr(strReport2(0)), App.Path & "\reports.ini")
    End If
    'If Left$(strdata(i), 9) = "USERRESET" Then
    '    Dim strReset As String
    '    strReset = Mid(strdata(i), 10, Len(strdata(i)))
    '    Dim iUserNum As Integer
    '    nSave = App.Path & "\user.ini"
    '    Debug.Print nSave
    '    Call WriteIni(strReset, "RATING", "1000", nfile)
    '    Call WriteIni(strReset, "WINS", "0", nfile)
    '    Call WriteIni(strReset, "LOSS", "0", nfile)
    '    Call WriteIni(strReset, "DISC", "0", nfile)
    '    Call WriteIni(strReset, "DJINNNUM", "1", nfile)
    '    For q = 1 To 20
    '        If Users(q).Name = strReset Then
    '            Chat(q).SendData "ADMINTXT" & "Your account has been reset by a chat moderator." & vbCrLf
    '            Chat(Index).SendData "ADMINTXT" & "Player's account has been reset." & vbCrLf
    '        End If
    '    Next 'q
    'End If

    If Left$(strData(i), 7) = "INFROWN" Then
        Call SendInChat("0", Index)
    End If
    If Left$(strData(i), 7) = "INSMILE" Then
        Call SendInChat("1", Index)
    End If
    If Left$(strData(i), 5) = "INDOT" Then
        Call SendInChat("2", Index)
    End If
    If Left$(strData(i), 3) = "IN!" Then
        Call SendInChat("3", Index)
    End If
    If Left$(strData(i), 3) = "IN?" Then
        Call SendInChat("4", Index)
    End If
    If Left$(strData(i), 7) = "INANGRY" Then
        Call SendInChat("6", Index)
    End If
    If Left$(strData(i), 6) = "INLOVE" Then
        Call SendInChat("5", Index)
    End If
    If Left$(strData(i), 6) = "INIDEA" Then
        Call SendInChat("7", Index)
    End If
    If Left$(strData(i), 7) = "INCLOUD" Then
        Call SendInChat("8", Index)
    End If
    
    
    If Left$(strData(i), 5) = "STATS" Then
        nfile = App.Path & "\user.ini"
        Dim strPlayer As String
        Dim curPlayer As Integer
        Dim strWins As String
        Dim strDisc As String
        Dim strLoss As String
        Dim strRate As String
        Dim strLev As String
        Dim strChr As String
        Dim strChr2 As String
        Dim strStatsCoins As String
        Dim strStatsItem As String
        Dim strStatsItem2 As String
        Dim strRanking As String
        strPlayer = Mid(strData(i), 6, Len(strData(i)))
        strWins = GetFromIni(strPlayer, "WINS", nfile)
        strLoss = GetFromIni(strPlayer, "LOSS", nfile)
        strDisc = GetFromIni(strPlayer, "DISC", nfile)
        strRate = GetFromIni(strPlayer, "RATING", nfile)
        strLev = GetFromIni(strPlayer, "LEVEL", nfile)
        strChr = GetFromIni(strPlayer, "CHAR", nfile)
        strChr2 = GetFromIni(strPlayer, "CHAR2", nfile)
        strStatsItem2 = GetFromIni(strPlayer, "ITEM2", nfile)
        strStatsCoins = GetFromIni(strPlayer, "COINS", nfile)
        strStatsItem = GetFromIni(strPlayer, "ITEM", nfile)
        strStatsItem = GetFromIni("I" & strStatsItem, "NAME", App.Path & "\items.ini")
        strStatsItem2 = GetFromIni("I" & strStatsItem2, "NAME", App.Path & "\items.ini")
        
        strRanking = GetFromIni(strPlayer, "RANK", nfile)
        If strRanking = "" Then strRanking = "Unranked"
        
        Chat(Index).SendData "CHATMSG" & strPlayer & "'s stats are: " & strWins & "/" & strLoss & "/" & strDisc & " at level " & strLev & " with a rating of " & strRate & " and " & strStatsCoins & " coins, using the characters " & strChr & " and " & strChr2 & ".  The player's current weapons are " & strStatsItem & " and " & strStatsItem2 & ".  The player's ranking in the ladder is: " & strRanking & vbCrLf
        DoEvents
    End If
    
    If Left$(strData(i), 7) = "SCRAMON" Then
        chkScrambler.Value = 1
        For q = 1 To 20
            If Users(q).Enabled = True Then
                Chat(q).SendData "SCRAMON" & vbCrLf
            End If
        Next 'q
    End If
    If Left$(strData(i), 8) = "SCRAMOFF" Then
        chkScrambler.Value = 0
        For q = 1 To 20
            If Users(q).Enabled = True Then
                Chat(q).SendData "SCRAMOFF" & vbCrLf
            End If
        Next 'q
    End If
    If Left$(strData(i), 8) = "NEEDCHAR" Then
        Dim sNeedChar As String
        sNeedChar = GetFromIni(Users(Index).Name, "CHAR", nfile)
        Chat(Index).SendData "HERECHAR" & sNeedChar & vbCrLf
    End If
    If Left$(strData(i), 4) = "MOTD" Then
        txtMOTD.Text = Mid$(strData(i), 5, Len(strData(i)))
        Call WriteIni("GEN", "MOTD", txtMOTD.Text, App.Path & "\motd.ini")
    End If
    If Left$(strData(i), 8) = "ANNOUNCE" Then
        Call SendDataToAll(strData(i))
    End If
    If Left$(strData(i), 7) = "SMODTXT" Then
        Call SendDataToAll(strData(i))
    End If
    If Left$(strData(i), 7) = "AWAYTXT" Then
        Call SendDataToAll(strData(i))
    End If
    If Left$(strData(i), 9) = "SERVERMSG" Then
        txtMsg.Text = Mid$(strData(i), 10, Len(strData(i)))
        Call cmdSend_Click
    End If
    If Left$(strData(i), 7) = "WINUSER" Then
    '    noclose = True
        nfile = App.Path & "\user.ini"
        curUser = Mid(strData(i), 8, Len(strData(i)))
        
        
    '    ServerNum(Index) = CInt(curUser)
        
        
    '    DoEvents
    
    '    Dim strTestName As String
    '    strTestName = GetFromIni(curuser, "NAME", nfile)
    '    If curUser <> strTestName And curUser <> "" Then
    '        ServerNum(Index) = FindUser(curUser)
    '    End If
        
    
    
    
    
        Call WriteIni(strTime, "USERNUM", curUser, App.Path & "\gamelog.ini")
        
    
    End If

    If Left$(strData(i), 6) = "RATING" Then
        curUser = Mid(strData(i), 7, Len(strData(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)
        strCurRating = GetFromIni(curUser, "RATING", nfile)
        iCurRating = CInt(strCurRating)
        Dim snewRating As String
        Dim inewRating As Integer
        snewRating = CStr(strsplit(1))
        inewRating = CInt(snewRating)
        
        Call WriteIni(strTime, "Gained Rating", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        
        snewRating = CStr(inewRating + iCurRating)
        iCurRating = inewRating + iCurRating
        iCurRating = iCurRating - 1000
        Dim intCurLvl As Long
        Dim intCurDjinn As Long
        intCurLvl = GetLevel(CStr(iCurRating))
        intCurDjinn = GetDjinn(CStr(iCurRating))
        Call WriteIni(curUser, "RATING", snewRating, nfile)
        Call WriteIni(curUser, "LEVEL", CStr(intCurLvl), nfile)
        Call WriteIni(curUser, "DJINNNUM", CStr(intCurDjinn), nfile)
    End If
    
    If Left$(strData(i), 5) = "COINS" Then
        curUser = Mid(strData(i), 6, Len(strData(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)
        
        Dim snewCoins As String
        Dim inewCoins As Integer
        strCurCoins = GetFromIni(curUser, "COINS", nfile)
        iCurCoins = CInt(strCurCoins)
        snewCoins = CStr(strsplit(1))
        
        
        Call WriteIni(strTime, "Gained Coins", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        
        inewCoins = CInt(snewCoins)
        snewCoins = CStr(inewCoins + iCurCoins)
        Call WriteIni(curUser, "COINS", snewCoins, nfile)
        
    End If
    'If Left$(strdata(i), 11) = "DJINNGAINED" Then
    '    curUser = Mid(strdata(i), 12, Len(strdata(i)))
    '    strsplit = Split(curUser, "@", -1, vbTextCompare)
    '    curUser = strsplit(0)

     '   Dim sDjinnGained As String
     '   Dim iDjinnGained As Integer
     '   sCurDjinn = GetFromIni(curUser, "DJINNNUM", nfile)
     '   iCurDjinn = CInt(sCurDjinn)
     '   sDjinnGained = CStr(strsplit(1))
     '   Call WriteIni(strTime, "DJINN-GAINED", curUser & sDjinnGained, App.Path & "\gamelog.ini")
     '   iDjinnGained = CInt(sDjinnGained)
     '   sDjinnGained = CStr(iDjinnGained + iCurDjinn)
     '   Call WriteIni(curUser, "DJINNNUM", sDjinnGained, nfile)
        
    'End If
    'If Left$(strdata(i), 3) = "LVL" Then
    '    curUser = Mid(strdata(i), 4, Len(strdata(i)))
    '    strsplit = Split(curUser, "@", -1, vbTextCompare)
    '    curUser = strsplit(0)
    '    Dim snewLVL As String
    '    Dim inewLVL As Integer
      '  strCurLvl = GetFromIni(curUser, "LEVEL", nfile)
      '  iCurLvl = CInt(strCurLvl)
      '  snewLVL = CStr(strsplit(1))
      '  Call WriteIni(strTime, "LVL-GAINED", curUser & snewLVL, App.Path & "\gamelog.ini")
      '  inewLVL = CInt(snewLVL)
       ' snewLVL = CStr(inewLVL + iCurLvl)
      '  Server(Index).SendData "STAT" & vbCrLf
     '   Call WriteIni(curUser, "LEVEL", snewLVL, nfile)

    'End If
    If Left$(strData(i), 4) = "SWIN" Then
        curUser = Mid(strData(i), 5, Len(strData(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)

        Call WriteIni(strTime, "Won Against", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        Dim CurWins As String
        Dim iWins As Integer
        CurWins = GetFromIni(curUser, "WINS", nfile)
        iWins = CInt(CurWins)
        CurWins = CStr(iWins + 1)
        Call WriteIni(curUser, "WINS", CurWins, nfile)
    End If
    If Left$(strData(i), 4) = "LOSE" Then
        curUser = Mid(strData(i), 5, Len(strData(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)

        Call WriteIni(strTime, "Lost Against", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        Dim CurLose As String
        Dim iLose As Integer
        CurLose = GetFromIni(curUser, "LOSS", nfile)
        iLose = CInt(CurLose)
        CurLose = CStr(iLose + 1)
        Call WriteIni(curUser, "LOSS", CurLose, nfile)
    End If
    
    If Left$(strData(i), 9) = "STATSLOSS" Then
        nfile = App.Path & "\user.ini"
        'Dim curUser As String
        curUser = Mid(strData(i), 10, Len(strData(i)))
        
        curMax = CInt(GetFromIni("GEN", "TOTAL", nfile))
        For q = 0 To curMax
            chkuser = GetFromIni(CStr(q), "NAME", nfile)
            If chkuser = curUser Then
            curUser = CStr(q)
            End If
            'DoEvents
        Next 'i
        
        strCurRating = GetFromIni(curUser, "RATING", nfile)
        iCurRating = CInt(strCurRating)
        strCurLvl = GetFromIni(curUser, "LEVEL", nfile)
        iCurLvl = CInt(strCurLvl)
        strCurCoins = GetFromIni(curUser, "COINS", nfile)
        iCurCoins = CInt(strCurCoins)
    End If
    
    If Left$(strData(i), 10) = "LOSSRATING" Then
            curUser = Mid(strData(i), 11, Len(strData(i)))
            strsplit = Split(curUser, "@", -1, vbTextCompare)
            curUser = strsplit(0)
            strCurRating = GetFromIni(curUser, "RATING", nfile)
            iCurRating = CInt(strCurRating)
    
            snewRating = CStr(strsplit(1))
            inewRating = CInt(snewRating)
            
            Call WriteIni(strTime, "Lost Rating", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
            
            
            snewRating = CStr(iCurRating - inewRating)
            iCurRating = iCurRating - inewRating
            iCurRating = iCurRating - 1000
            If iCurRating < 0 Then iCurRating = 0
    
            intCurLvl = GetLevel(CStr(iCurRating))
            intCurDjinn = GetDjinn(CStr(iCurRating))
            Call WriteIni(curUser, "RATING", snewRating, nfile)
            Call WriteIni(curUser, "LEVEL", CStr(intCurLvl), nfile)
            Call WriteIni(curUser, "DJINNNUM", CStr(intCurDjinn), nfile)
        
    End If
    
    If Left$(strData(i), 9) = "LOSSCOINS" Then
            curUser = Mid(strData(i), 10, Len(strData(i)))
            strsplit = Split(curUser, "@", -1, vbTextCompare)
            curUser = strsplit(0)
    
            strCurCoins = GetFromIni(curUser, "COINS", nfile)
            iCurCoins = CInt(strCurCoins)
            snewCoins = CStr(strsplit(1))
            
            
            Call WriteIni(strTime, "Lost Coins", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
            
            
            inewCoins = CInt(snewCoins)
            snewCoins = CStr(iCurCoins - inewCoins)
            Call WriteIni(curUser, "COINS", snewCoins, nfile)
    End If

Next 'i

    If strdatao <> "CHATSPAM" & vbCrLf And DontAdd = False Then
        lstStatus.AddItem strTime & " " & strdatao
    End If
    
Exit Sub
err:
Debug.Print err.Description
Exit Sub


End Sub

Private Sub chkScrambler_Click()
On Error Resume Next
    If chkScrambler.Value = 1 Then
    Call AdminChat("Scrambler enabled by the server host.")
        For q = 1 To 20
            If Users(q).Enabled = True Then
                Chat(q).SendData "SCRAMON" & vbCrLf

            End If
            DoEvents
        Next 'q
    Else
    Call AdminChat("Scrambler disabled by the server host.")
        For q = 1 To 20
            If Users(q).Enabled = True Then
                Chat(q).SendData "SCRAMOFF" & vbCrLf
            End If
        DoEvents
        Next 'q
    End If

End Sub

Private Sub closeme_Timer()
On Error GoTo err
For i = 1 To 20
    If Server(i).State <> sckClosed Then
        Server(i).SendData "P" & vbCrLf
    End If
    DoEvents
Next 'i
Exit Sub
err:
    Server(i).Close
    Resume Next
End Sub

Private Sub cmdAdd_Click(Index As Integer)
On Error Resume Next
Dim nSave As String
Dim itotal As Integer
Dim sTotal As String
If Index = 0 Then
nSave = App.Path & "\psynergy.ini"
    If txtPsy(1).Text = "W" Then
        sTotal = GetFromIni("GEN", "W", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "W", sTotal, nSave)
        sTotal = "W" & CStr(itotal)
    End If
    If txtPsy(1).Text = "F" Then
        sTotal = GetFromIni("GEN", "F", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "F", sTotal, nSave)
        sTotal = "F" & CStr(itotal)
    End If
    If txtPsy(1).Text = "N" Then
        sTotal = GetFromIni("GEN", "N", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "N", sTotal, nSave)
        sTotal = "N" & CStr(itotal)
    End If
    If txtPsy(1).Text = "E" Then
        sTotal = GetFromIni("GEN", "E", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "E", sTotal, nSave)
        sTotal = "E" & CStr(itotal)
    End If
    If txtPsy(1).Text = "D" Then
        sTotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "D", sTotal, nSave)
        sTotal = "D" & CStr(itotal)
    End If
    If txtPsy(1).Text = "H" Then
        sTotal = GetFromIni("GEN", "H", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "H", sTotal, nSave)
        sTotal = "H" & CStr(itotal)
    End If
    Call WriteIni(sTotal, "NAME", txtPsy(0).Text, nSave)
    Call WriteIni(sTotal, "TYPE", txtPsy(2).Text, nSave)
    Call WriteIni(sTotal, "DAMAGE", txtPsy(3).Text, nSave)
    Call WriteIni(sTotal, "PP", txtPsy(4).Text, nSave)
    Call WriteIni(sTotal, "RATING", txtPsy(5).Text, nSave)
    Call WriteIni(sTotal, "DJINN", txtPsy(6).Text, nSave)
    Call WriteIni(sTotal, "DESC", txtPsy(7).Text, nSave)
End If
If Index = 1 Then
nSave = App.Path & "\items.ini"
    sTotal = GetFromIni("GEN", "TOTAL", nSave)
    itotal = CInt(sTotal)
    itotal = itotal + 1
    sTotal = CStr(itotal)
    Call WriteIni("GEN", "TOTAL", sTotal, nSave)
    sTotal = CStr(itotal)
    Call WriteIni("I" & sTotal, "NAME", txtItem(0).Text, nSave)
    Call WriteIni("I" & sTotal, "COINS", txtItem(1).Text, nSave)
    Call WriteIni("I" & sTotal, "DESCRIPTION", txtItem(2).Text, nSave)
    Call WriteIni("I" & sTotal, "SPCNAME", txtItem(3).Text, nSave)
    Call WriteIni("I" & sTotal, "SPCDAMAGE", txtItem(5).Text, nSave)
    Call WriteIni("I" & sTotal, "TYPE", txtItem(6).Text, nSave)
    Call WriteIni("I" & sTotal, "DAMAGE", txtItem(7).Text, nSave)
    Call WriteIni("I" & sTotal, "ADDMOD", txtItem(8).Text, nSave)
    Call WriteIni("I" & sTotal, "MULTMOD", txtItem(9).Text, nSave)
    Call WriteIni("I" & sTotal, "SPCPERCENT", txtItem(10).Text, nSave)

End If
If Index = 2 Then
nSave = App.Path & "\djinn.ini"
    If txtDjinn(1).Text = "W" Then
        sTotal = GetFromIni("GEN", "W", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "W", sTotal, nSave)
        sTotal = "W" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "F" Then
        sTotal = GetFromIni("GEN", "F", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "F", sTotal, nSave)
        sTotal = "F" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "N" Then
        sTotal = GetFromIni("GEN", "N", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "N", sTotal, nSave)
        sTotal = "N" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "E" Then
        sTotal = GetFromIni("GEN", "E", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "E", sTotal, nSave)
        sTotal = "E" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "D" Then
        sTotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "D", sTotal, nSave)
        sTotal = "D" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "H" Then
        sTotal = GetFromIni("GEN", "H", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "H", sTotal, nSave)
        sTotal = "H" & CStr(itotal)
    End If
    Call WriteIni(sTotal, "NAME", txtDjinn(0).Text, nSave)
    Call WriteIni(sTotal, "DESCRIPTION", txtDjinn(2).Text, nSave)
    Call WriteIni(sTotal, "TYPE", txtDjinn(3).Text, nSave)
    Call WriteIni(sTotal, "DAMAGE", txtDjinn(4).Text, nSave)
    Call WriteIni(sTotal, "ADDMOD", txtDjinn(5).Text, nSave)
End If
If Index = 3 Then
    nSave = App.Path & "\summons.ini"
    If txtSummon(2).Text = "W" Then
        sTotal = GetFromIni("GEN", "W", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "W", sTotal, nSave)
        sTotal = "W" & CStr(itotal)
    End If
    If txtSummon(2).Text = "F" Then
        sTotal = GetFromIni("GEN", "F", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "F", sTotal, nSave)
        sTotal = "F" & CStr(itotal)
    End If
    If txtSummon(2).Text = "N" Then
        sTotal = GetFromIni("GEN", "N", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "N", sTotal, nSave)
        sTotal = "N" & CStr(itotal)
    End If
    If txtSummon(2).Text = "E" Then
        sTotal = GetFromIni("GEN", "E", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "E", sTotal, nSave)
        sTotal = "E" & CStr(itotal)
    End If
    If txtSummon(2).Text = "D" Then
        sTotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "D", sTotal, nSave)
        sTotal = "D" & CStr(itotal)
    End If
    If txtSummon(2).Text = "H" Then
        sTotal = GetFromIni("GEN", "H", nSave)
        itotal = CInt(sTotal)
        itotal = itotal + 1
        sTotal = CStr(itotal)
        Call WriteIni("GEN", "H", sTotal, nSave)
        sTotal = "H" & CStr(itotal)
    End If
    Call WriteIni(sTotal, "NAME", txtSummon(0).Text, nSave)
    Call WriteIni(sTotal, "DESC", txtSummon(1).Text, nSave)
    Call WriteIni(sTotal, "DJINN", txtSummon(3).Text, nSave)
End If
    
End Sub

Private Sub cmdAnnounce_Click()
On Error Resume Next
For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "ANNOUNCE" & txtMsg.Text & vbCrLf
    End If
DoEvents
Next 'q

txtChat.Text = txtChat.Text & vbNewLine & "Server: " & txtMsg.Text
txtMsg.Text = ""
End Sub

Private Sub cmdCloseServer_Click()
On Error Resume Next
If cmdCloseServer.Caption <> "Re-open Server" Then
    cmdCloseServer.Caption = "Re-open Server"
    strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
        txtHTML.Text = ""
        txtHTML.Text = "<b>The server has been temporarily shut down for paitent.<br><br>Current Work on the Game: <b>" & txtMOTD.Text & "</b>"
 FNO = FreeFile
     On Error Resume Next
     err.Clear
     Open ("C:\sambar52\docs\status.html") For Output As #FNO
      If err.Number <> 0 Then
        'MsgBox "an error has occured"
       Else
        Print #FNO, (txtHTML.Text)
      End If
     Close #FNO
     'On Error GoTo 0
Call WriteIni("GEN", "DOWN", "TRUE", App.Path & "\motd.ini")
Else
    cmdCloseServer.Caption = "Go Down For Maitence"
Call WriteIni("GEN", "DOWN", "FALSE", App.Path & "\motd.ini")
End If
End Sub

Private Sub cmdCustSave_Click()
On Error Resume Next
Dim intCustMax As Long
Dim strCustMax As String
Dim cSave As String
cSave = App.Path & "\customchar.ini"

strCustMax = GetFromIni("GEN", "TOTAL", cSave)
If strCustMax = "" Then strCustMax = "0"
intCustMax = CInt(strCustMax)
intCustMax = intCustMax + 1
strCustMax = CStr(intCustMax)
Call WriteIni("GEN", "TOTAL", strCustMax, cSave)

Call WriteIni(strCustMax, "NAME", txtCustChar(0).Text, cSave)
Call WriteIni(strCustMax, "CLASS", txtCustChar(1).Text, cSave)
Call WriteIni(strCustMax, "HP", txtCustChar(2).Text, cSave)
Call WriteIni(strCustMax, "PP", txtCustChar(3).Text, cSave)
Call WriteIni(strCustMax, "AP", txtCustChar(4).Text, cSave)
Call WriteIni(strCustMax, "DEFENSE", txtCustChar(5).Text, cSave)
Call WriteIni(strCustMax, "POWER", txtCustChar(6).Text, cSave)
Call WriteIni(strCustMax, "RESIST", txtCustChar(7).Text, cSave)
Call WriteIni(strCustMax, "DESCRIPTION", txtCustChar(8).Text, cSave)
Call WriteIni(strCustMax, "USERS", txtCustChar(9).Text, cSave)
Call WriteIni(strCustMax, "LUCK", txtCustChar(12).Text, cSave)
Call WriteIni(strCustMax, "STRENGTH", txtCustChar(10).Text, cSave)
Call WriteIni(strCustMax, "WEAKNESS", txtCustChar(11).Text, cSave)
Call WriteIni(strCustMax, "PICTURE", txtCustChar(13).Text, cSave)


For i = 0 To txtCustChar.UBound
    txtCustChar(i).Text = ""
Next 'i

End Sub

Private Sub cmdGame_Click()
On Error Resume Next
frmEditor.Show
End Sub

Private Sub cmdKill_Click()
On Error Resume Next
For q = 1 To 20
If Users(q).Name = txtKill.Text Then
    Chat(q).SendData "CHATKILL" & vbCrLf
    Chat(q).Close
    Users(q).Enabled = False
End If
DoEvents
Next 'q

End Sub

Private Sub cmdMOTD_Click()
On Error Resume Next
Call WriteIni("MOTD", "MOTD", txtMOTD.Text, App.Path & "\motd.ini")
End Sub

Private Sub cmdMsg_Click()
On Error Resume Next
For q = 1 To 20
If Users(q).Name = txtKill.Text Then
Chat(q).SendData "CHATTXT" & "[PRIVATE MESSAGE FROM SERVER]: " & txtContent.Text & vbCrLf
DoEvents
End If
Next 'q
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
filChar.Refresh
filChar.Pattern = "*.omaz;*.gif"
End Sub

Private Sub cmdReset_Click()
On Error Resume Next
yadda = MsgBox("Are you sure?", vbYesNo)
If yadda <> vbYes Then
    Exit Sub
End If

Dim strMax As String
Dim iMax As Integer
Dim nSave As String
Dim strCheckWins As String
Dim strCheckLoss As String
Dim strCheckRating As String
Dim strCheckKarma As String
Dim realMax As Integer
Dim newSave As String
newSave = App.Path & "\newuser.ini"

realMax = 0

nSave = App.Path & "\user.ini"
strMax = GetFromIni("GEN", "TOTAL", nSave)

vbinput = MsgBox("Reset stats?", vbYesNo)

iMax = CInt(strMax)
    Call AdminChat("[Server Message:] Ladder reset has begun.")
For i = 0 To iMax
    Dim resetUser As String
    resetUser = GetFromIni("USERNUM", CStr(i), nSave)
    strCheckWins = GetFromIni(resetUser, "WINS", nSave)
    strCheckLoss = GetFromIni(resetUser, "LOSS", nSave)
    strCheckRating = GetFromIni(resetUser, "RATING", nSave)
    strCheckKarma = GetFromIni(resetUser, "KARMA", nSave)
    If strCheckKarma = "" Then strCheckKarma = "0"
    If (strCheckWins <> "0" And strCheckWins <> "") Or (strCheckLoss <> "0" And strCheckLoss <> "") Or (strCheckRating <> "1000" And strCheckRating <> "") Or CInt(strCheckKarma) > 10 Or resetUser = "dragoon" Or resetUser = "tacvek" Or resetUser = "admin" Then 'If the account is active
        realMax = realMax + 1
        Dim strTransfer As String
        
        If vbinput = vbYes Then
            Call WriteIni("USERNUM", CStr(realMax), resetUser, newSave)
            Call WriteIni(resetUser, "RATING", "1000", newSave)
            Call WriteIni(resetUser, "LEVEL", "1", newSave)
            Call WriteIni(resetUser, "WINS", "0", newSave)
            Call WriteIni(resetUser, "LOSS", "0", newSave)
            Call WriteIni(resetUser, "DISC", "0", newSave)
            Call WriteIni(resetUser, "COINS", "100", newSave)
            Call WriteIni(resetUser, "ITEM", "1", newSave)
            Call WriteIni(resetUser, "DJINNNUM", "1", newSave)
        Else
            Call WriteIni("USERNUM", CStr(realMax), resetUser, newSave)
            strTransfer = GetFromIni(resetUser, "RATING", nSave)
            Call WriteIni(resetUser, "RATING", strTransfer, newSave)
            strTransfer = GetFromIni(resetUser, "LEVEL", nSave)
            Call WriteIni(resetUser, "LEVEL", strTransfer, newSave)
            strTransfer = GetFromIni(resetUser, "WINS", nSave)
            Call WriteIni(resetUser, "WINS", strTransfer, newSave)
            strTransfer = GetFromIni(resetUser, "LOSS", nSave)
            Call WriteIni(resetUser, "LOSS", strTransfer, newSave)
            strTransfer = GetFromIni(resetUser, "DISC", nSave)
            Call WriteIni(resetUser, "DISC", strTransfer, newSave)
            strTransfer = GetFromIni(resetUser, "COINS", nSave)
            Call WriteIni(resetUser, "COINS", strTransfer, newSave)
            strTransfer = GetFromIni(resetUser, "ITEM", nSave)
            Call WriteIni(resetUser, "ITEM", strTransfer, newSave)
            strTransfer = GetFromIni(resetUser, "DJINNNUM", nSave)
            Call WriteIni(resetUser, "DJINNNUM", strTransfer, newSave)
        End If
        
        strTransfer = GetFromIni(resetUser, "PASSWORD", nSave)
        Call WriteIni(resetUser, "PASSWORD", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "CHAR", nSave)
        Call WriteIni(resetUser, "CHAR", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "TYPE", nSave)
        Call WriteIni(resetUser, "TYPE", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PASSWORD", nSave)
        Call WriteIni(resetUser, "PASSWORD", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "IP", nSave)
        Call WriteIni(resetUser, "IP", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PASSWORD", nSave)
        Call WriteIni(resetUser, "PASSWORD", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PNAME", nSave)
        Call WriteIni(resetUser, "PNAME", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PAGE", nSave)
        Call WriteIni(resetUser, "PAGE", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PLOCATION", nSave)
        Call WriteIni(resetUser, "PLOCATION", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PSEX", nSave)
        Call WriteIni(resetUser, "PSEX", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PAIM", nSave)
        Call WriteIni(resetUser, "PAIM", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PEMAIL", nSave)
        Call WriteIni(resetUser, "PEMAIL", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PMSN", nSave)
        Call WriteIni(resetUser, "PMSN", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "POTHER", nSave)
        Call WriteIni(resetUser, "POTHER", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "KARMA", nSave)
        Call WriteIni(resetUser, "KARMA", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "LASTON", nSave)
        Call WriteIni(resetUser, "LASTON", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PICQ", nSave)
        Call WriteIni(resetUser, "PICQ", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "AVATAR", nSave)
        Call WriteIni(resetUser, "AVATAR", strTransfer, newSave)
        strTransfer = GetFromIni(resetUser, "PINNUM", nSave)
        Call WriteIni(resetUser, "PINNUM", strTransfer, newSave)
    End If
    DoEvents
Next 'i
    Call AdminChat("[Server Message:] Ladder reset has finished.")

End Sub

Private Sub cmdSend_Click()
On Error Resume Next
If txtMsg.Text <> "close" Then
For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "ADMINTXT" & "[SERVER MESSAGE]: " & txtMsg.Text & vbCrLf
    End If
DoEvents
Next 'q
Else
Server(0).Close
Server(0).Listen
End If

txtChat.Text = txtChat.Text & vbNewLine & "Server: " & txtMsg.Text
txtMsg.Text = ""

End Sub

Private Sub cmdShowCustChar_Click()
On Error Resume Next
If framCustomChar.Visible = True Then
    framCustomChar.Visible = False
Else
    framCustomChar.Visible = True
    framCustomChar.Width = Me.ScaleWidth - framCustomChar.Left
    framCustomChar.Height = txtChat.Top
End If
End Sub

Private Sub cmdTOS_Click()
On Error Resume Next
Dim iTos As Integer
Dim nSave As String

nSave = App.Path & "\motd.ini"
strSambarPath = GetFromIni("GEN", "SAMBAR", App.Path & "\motd.ini")
intDjinnSaveHighScore = CInt(GetFromIni("GEN", "SCORE", App.Path & "\motd.ini"))
strDjinnSavePlayer = GetFromIni("GEN", "PLAYER", App.Path & "\motd.ini")


iTos = CInt(GetFromIni("MOTD", "TOSMAX", nSave))
For i = 1 To iTos
    strTOS(i) = GetFromIni("MOTD", "TOS" & i, nSave)
    DoEvents
Next 'i
End Sub

Private Sub cmdUpdateLadder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call UpdateLadder
Else
    strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
        txtHTML.Text = ""
        Dim intTotal As Integer
        For i = 1 To 20
            If Users(i).Enabled = True Then
                intTotal = intTotal + 1
            End If
            DoEvents
        Next 'i
        strtotal = CStr(intTotal)
        struptime = CStr(upTime)
        strLastUser = GetFromIni("GEN", "NEWEST", App.Path & "\user.ini")
        Dim strScrambler As String
        If chkScrambler.Value = 1 Then
            strScrambler = "On"
        Else
            strScrambler = "Off"
        End If
        txtHTML.Text = "This webpage is updated about every 15 minutes if the server is up.  If it hasn't been updated within the last few minutes, then the server is probably down for maitence.  Please don't post issues about the server being down on the Bug Report forum.<br><br>As of <b>" & strTime & "</b> the server is running <b>Version " & txtVer.Text & "</b> and has been running for <b>" & CStr(CInt(struptime) * 15) & "</b> minutes.<br><br>There are currently <b>" & strtotal & "</b> users online.<br>The last user to register was <b>" & strLastUser & "</b><br>Scrambler is currently: <b>" & strScrambler & "</b><br><br>Message of the Day: <b>" & txtMOTD.Text & "</b>"
 FNO = FreeFile
     On Error Resume Next
     err.Clear
     Open (strSambarPath & "\status.html") For Output As #FNO
      If err.Number <> 0 Then
        'MsgBox "an error has occured"
       Else
        Print #FNO, (txtHTML.Text)
      End If
     Close #FNO
     'On Error GoTo 0
    End If
End Sub

Private Sub FileTransfer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
If Index = 0 Then
    'Beep
    iNewServer = 100
    For q = 1 To 7
    If FileTransfer(q).State = sckClosed And iNewServer = 100 Then
        FileTransfer(q).Accept requestID

        DoEvents

        Call SendList(q, False)
            
        Exit Sub
    End If
    Next 'i
End If 'if index = 0
End Sub

Private Sub FileTransfer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
FileTransfer(Index).GetData arrdatao
arrdata = Split(arrdatao, vbCrLf, -1, vbTextCompare)

For i = 0 To UBound(arrdata)

    'If Left$(arrdata(i), 7) = "RQSTLST" Then        'if user requested file list
    '        strTemp = ""
    '        strFileList = ""
    '        filChar.Path = App.Path & "\files\"  'get the first filename
    '        If filChar.List(0) = "" Then                    'if there are no files then
    '            FileTransfer(Index).SendData "LSTnul"    'send the nofiles message (nul is a virual file that exists in every directory)
    '        Else                                    'else if there are files then start collecting a list of all files
    '            Call FileListLoop(0, Index, False)
    '        End If
    'End If
    If Left$(arrdata(i), 6) = "RQSTFL" Then        'if user requested file
        Dim strFile As String
        Dim strNewFile As String
        Dim strEncodedFile As String
        strFile = Mid$(arrdata(i), 7, Len(arrdata(i)))
        If Right$(strFile, 3) <> "ini" Then
            strNewFile = App.Path & "\files\" & strFile
        Else
            strNewFile = App.Path & "\" & strFile
        End If
        Dim b64 As New base64
        strEncodedFile = b64.EncodeFromFile(strNewFile)
        FileTransfer(Index).SendData "FILE" & strFile & "@" & strEncodedFile & "!" 'send the file
    ElseIf Left$(arrdata(i), 5) = "CLOSE" Then
        FileTransfer(Index).Close
        GettingFile = False
    ElseIf Left$(arrdata(i), 5) = "CLEAR" Then
        GettingFile = False
        File = ""
    Else
        If GettingFile = False Then
            GettingFile = True
            File = arrdata(i)
        Else
            File = File & arrdata(i)
        End If
        If InStr(arrdata(i), "!") Then
            GettingFile = False
            Call MakeFile(File)
            File = ""
            txtMsg.Text = "Recieved Maze File Successfully!"
            Call cmdSend_Click
'            Call DownloadNextFile
        End If
    End If
Next 'i



End Sub

Private Sub Form_DblClick()
On Error Resume Next
If frmServer(Index).Height > 1200 Then
frmServer(Index).Height = 1200
frmServer(Index).Width = 1200
Else
frmServer(Index).Height = 7830
frmServer(Index).Width = 7830
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
GettingFile = False
filChar.Pattern = "*.gif;*"
filChar.Refresh
Server(0).Listen
Chat(0).Listen
FileTransfer(0).Listen
GettingFile = False

bCurUpdating = False
loginWait = False
lblIP.Caption = Server(0).LocalIP


Call cmdTOS_Click

ladderRefresh = 0

ReDim Users(21)
ReDim Game(21)

For i = 1 To 20
Load Server(i)
Load Chat(i)
Users(i).Enabled = False
Users(i).Away = False
Users(i).Number = 0
Game(i).Enabled = False
DoEvents

If i <= 7 Then
    Load FileTransfer(i)
    FileTransfer(i).LocalPort = 9880
End If

Next 'i
Load Server(21)

txtMOTD.Text = GetFromIni("MOTD", "MOTD", App.Path & "\motd.ini")
txtVer.Text = GetFromIni("MOTD", "VER", App.Path & "\motd.ini")
intDjinnSaveHighScore = CInt(GetFromIni("GEN", "SCORE", App.Path & "\motd.ini"))
strSambarPath = GetFromIni("GEN", "SAMBAR", App.Path & "\motd.ini")
strDjinnSavePlayer = GetFromIni("GEN", "PLAYER", App.Path & "\motd.ini")

ServerVersion = txtVer.Text
strdown = GetFromIni("GEN", "DOWN", App.Path & "\motd.ini")
If strdown = "FALSE" Then
    cmdCloseServer.Caption = "Close Server For Maitence"
Else
    cmdCloseServer.Caption = "Re-Open Server"
End If

upTime = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
'frmEditor.Show
End If
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
For i = 1 To 20
    Chat(i).SendData "DISC" & vbCrLf
    DoEvents
Next 'i
End
End Sub

Private Sub lstStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    yadda = MsgBox("Clear?", vbYesNo)
    If yadda = vbYes Then
        lstStatus.Clear
    End If
End If
End Sub

Private Sub nChat_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'On Error GoTo err
'Change for the ladder tournament:
Dim userSave As String
Dim dataSave As String
userSave = App.Path & "\user.ini"
dataSave = App.Path & "\serverdata.ini"
Dim strShortTime As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
strShortTime = Format(Now, "dd-mmmm")
'strTime = CStr(curTime)
Dim strRawData As String
nChat(Index).GetData strRawData
strData = Split(strRawData, vbCrLf, -1, vbTextCompare)


Exit Sub
'err:
'Resume Next
End Sub

Private Sub Server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'On Error GoTo err
If Index = 0 Then
    'Beep
    iNewServer = 100
    Dim intBanMax As String
    Dim bSave As String
    Dim strCheckBan As String
    Dim bBan As Boolean
    bSave = App.Path & "\ban.ini"

        Call WriteIni("GEN", "MAX", CStr(0), bSave)
    intBanMax = CInt(GetFromIni("GEN", "MAX", bSave))
    bBan = False
    For i = 1 To intBanMax
        strCheckBan = GetFromIni("GEN", CStr(i), bSave)
        If Mid$(strCheckBan, Len(strCheckBan), 1) = "*" Then
            If Left$(strCheckBan, Len(strCheckBan) - 1) = Left$((Server(Index).RemoteHostIP), Len(strCheckBan) - 1) Then
                bBan = True
            End If
        Else
            If strCheckBan = Server(Index).RemoteHostIP Then
                bBan = True
            End If
        End If
    Next 'i
    
    For q = 1 To 20
    
    If Server(q).State = sckClosed And iNewServer = 100 Then
        If bBan = False Then
            Server(q).Accept requestID
            For i = 1 To 25
                If strTOS(i) <> "" Then
                    Server(q).SendData "TOS" & strTOS(i) & vbCrLf
                End If
                DoEvents
            Next 'i
            iNewServer = MaxCon
            Exit Sub
        Else
            'Nohting
        End If
    End If
    Next 'i
    If iNewServer = 100 And MaxCon <> 100 Then
        Server(21).Accept requestID
        Server(21).SendData "FULL" & vbCrLf
        DoEvents
        Server(21).Close
    End If
End If 'if index = 0
Exit Sub
'err:
Exit Sub

End Sub

Private Sub Server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo err
    Dim nfile As String
    Dim xfile As String
    'Comment out when not in Ladder Tournament:
    nfile = App.Path & "\user.ini"
    xfile = App.Path & "\data.ini"
Dim strtime2 As String
Dim strGameTime As String

strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
strtime2 = Format(Now, "dd-mmmm")
strGameTime = Format(Now, "dd-mmmm hh:mm AM/PM")
'strTime = CStr(curTime)

Server(Index).GetData arrdatao
arrdata = Split(arrdatao, vbCrLf, -1, vbTextCompare)

For i = 0 To UBound(arrdata)
'If i >= UBound(arrdata) Then Exit Sub

If Left$(arrdata(i), 10) = "CHANGEUSER" Then
    strChangeUser = Mid$(arrdata(i), 11, Len(arrdata(i)))
End If
If Left$(arrdata(i), 11) = "CHANGEOLDPW" Then
    strChangePass = Mid$(arrdata(i), 12, Len(arrdata(i)))
End If
If Left$(arrdata(i), 11) = "CHANGENEWPW" Then
    strNewPass = Mid$(arrdata(i), 12, Len(arrdata(i)))
    strChangeRealPW = GetFromIni(strChangeUser, "PASSWORD", App.Path & "\user.ini")
    If strChangeRealPW = strChangePass Then
        Call WriteIni(strChangeUser, "PASSWORD", strNewPass, App.Path & "\user.ini")
        Server(Index).SendData "CHANGEPWGOOD" & vbCrLf
    Else
        Server(Index).SendData "CHANGEPWBAD" & vbCrLf
    End If
End If

If Left$(arrdata(i), 4) = "USER" Then
    UserVersion = 0
    curUser = Mid(arrdata(i), 5, Len(arrdata(i)))
    Debug.Print "USER" & " " & curUser
    'MsgBox curUser
    noclose = True
    'Exit Sub
    If curUser = "" Then Server(Index).Close
    For q = 1 To 20
        If Users(q).Name = curUser And Users(q).Enabled = True Then
            Server(Index).Close
            Exit Sub
        End If
    Next 'i
End If
If Left$(arrdata(i), 4) = "VERS" Then
    Dim strVersion As String
    strVersion = Mid(arrdata(i), 5, Len(arrdata(i)))
    UserVersion = CVar(strVersion)
    If strVersion <> ServerVersion Then
        Server(Index).SendData "VERSBAD" & vbCrLf
    End If
End If
If Left$(arrdata(i), 6) = "PINNUM" Then
    strNewPin = Mid(arrdata(i), 7, Len(arrdata(i)))
End If

If Left$(arrdata(i), 4) = "PASS" And loginWait = False Then
    Server(Index).SendData "DATE" & strtime2 & vbCrLf
    Debug.Print "PASS"
    ladUser = False
    loginWait = True
    timeWait.Enabled = False
    timeWait.Enabled = True
    'Comment out for Ladder Tournament:
    
    nfile = App.Path & "\user.ini"
    curPassword = Mid(arrdata(i), 5, Len(arrdata(i)))
        
    Dim bSwear As Boolean
    bSwear = CheckSwear(curUser)
    If curPassword = "alcatraz" Or curPassword = "celibi" Or curUser = "psibolt" Or curUser = "zidane" Then
        Dim bSave As String
        Dim iBanMax As Integer
        bSave = App.Path & "\ban.ini"
        iBanMax = GetFromIni("GEN", "MAX", bSave)
        Call WriteIni("GEN", CStr(iBanMax + 1), Server(Index).RemoteHostIP, bSave)
        Call WriteIni("GEN", "MAX", CStr(iBanMax + 1), bSave)
            'Server(Index).SendData "TOSlol, ur banned guylol, ur banned guylol, ur banned guylol, ur banned guylol, ur banned guylol, ur banned guylol, ur banned guylol, ur banned guylol, ur banned guy"
        Server(Index).Close
        Call WriteIni(curUser, "Log On", strTime, App.Path & "\zidane.ini")
    End If
    
    Dim intPinBanMax As Integer
    Dim strTempPin As String
    Dim strCurPin As String
    intPinBanMax = CInt(GetFromIni("GEN", "TOTAL", App.Path & "\pinban.ini"))
    strCurPin = GetFromIni(curUser, "PINNUM", nfile)
    For q = 1 To intPinBanMax
        If strCurPin = "" Then strCurPin = "1"
        strTempPin = GetFromIni("GEN", CStr(q), App.Path & "\pinban.ini")
        If strTempPin = strCurPin And strCurPin <> "" Then
            Server(Index).SendData "BADPIN" & vbCrLf
            DoEvents
            Server(Index).SendData "WINS" & "0" & vbCrLf
            Server(Index).SendData "LOSS" & "9001" & vbCrLf
            Server(Index).SendData "DISC" & "0" & vbCrLf
            Server(Index).SendData "RATING" & "-9001" & vbCrLf
            DoEvents
            Server(Index).SendData "FULL" & vbCrLf
            DoEvents
            Server(Index).Close
            Exit Sub
        End If
        DoEvents
    Next 'q

    Dim curMax As Integer
    curMax = CInt(GetFromIni("GEN", "TOTAL", nfile))
    
    Dim curuser2 As String
    
'    Dim iCurUser As Integer
'    iCurUser = FindUser(curUser)
'    ServerNum(Index) = iCurUser
'    Debug.Print ServerNum(Index)

'    curUser = CStr(iCurUser)
'    Debug.Print curUser

       
    
'    Server(Index).SendData "NUMUSER" & curUser & vbCrLf
    

    realPassword = GetFromIni(curUser, "PASSWORD", nfile)
    strRating = GetFromIni(curUser, "RATING", nfile)
    strChar(1) = GetFromIni(curUser, "CHAR", nfile)
    strChar(2) = GetFromIni(curUser, "CHAR2", nfile)
    If strChar(2) = "" Then
        strChar(2) = "Isaac"
        Call WriteIni(curUser, "CHAR2", "Isaac", nfile)
        Call WriteIni(curUser, "TYPE2", "E", nfile)
        Call WriteIni(curUser, "ITEM2", "1", nfile)
    End If
    strLvl = GetFromIni(curUser, "LEVEL", nfile)
    strDjinn = GetFromIni(curUser, "DJINNNUM", nfile)
    strCoins = GetFromIni(curUser, "COINS", nfile)
    strWins = GetFromIni(curUser, "WINS", nfile)
    strDisc = GetFromIni(curUser, "DISC", nfile)
    strLoss = GetFromIni(curUser, "LOSS", nfile)
    strType(1) = GetFromIni(curUser, "TYPE", nfile)
    strType(2) = GetFromIni(curUser, "TYPE2", nfile)
    strMyWeapon(1) = GetFromIni(curUser, "ITEM", nfile)
    strMyWeapon(2) = GetFromIni(curUser, "ITEM2", nfile)
    
    Users(Index).Rating = strRating
    Users(Index).Wins = strWins
    Users(Index).Losses = strLoss
    Users(Index).Disconnects = strDisc
    
    
    intLastDjinn = 0
    intLastPsy = 0
    intLastSummon = 0
    
    'MsgBox curPassword
    If curPassword = realPassword And curPassword <> "" And UserVersion = ServerVersion Then
        Call WriteIni(strTime, "Data", "Good Password Attempt on " & curUser & " at " & Server(Index).RemoteHostIP, xfile)
        Call WriteIni(curUser, "IP", Server(Index).RemoteHostIP, nfile)
        
        Call WriteIni(curUser, "PINNUM", strNewPin, nfile)
        
        Dim modSave As String
        Dim intModTotal As Integer
        Dim strTempMod As String
        modSave = App.Path & "\mods.ini"
        intModTotal = CInt(GetFromIni("GEN", "TOTAL", modSave))
        For q = 1 To intModTotal
            strTempMod = GetFromIni(CStr(q), "NAME", modSave)
            Server(Index).SendData "MODNAME" & strTempMod & vbCrLf
            DoEvents
        Next 'q
        
        
        Dim nSave As String
        Dim sTotal As String
        Dim itotal As Integer
        nSave = App.Path & "\items.ini"
        sTotal = GetFromIni("GEN", "TOTAL", nSave)
        itotal = CInt(sTotal)
        Server(Index).SendData "HIGHSCOREDS" & intDjinnSaveHighScore & vbCrLf
        DoEvents
        Server(Index).SendData "HIGHPLAYERDS" & strDjinnSavePlayer & vbCrLf
        DoEvents
        Server(Index).SendData "RATING" & strRating & vbCrLf
        DoEvents
        Server(Index).SendData "WPN" & strMyWeapon(1) & vbCrLf
        DoEvents
        Server(Index).SendData "2WPN" & strMyWeapon(2) & vbCrLf
        DoEvents
        Server(Index).SendData "CHAR" & strChar(1) & vbCrLf
        DoEvents
        Server(Index).SendData "2CHAR" & strChar(2) & vbCrLf
        DoEvents
        Server(Index).SendData "DJINN" & strDjinn & vbCrLf
        DoEvents
        Server(Index).SendData "COINS" & strCoins & vbCrLf
        DoEvents
        Server(Index).SendData "WINS" & strWins & vbCrLf
        DoEvents
        Server(Index).SendData "LOSS" & strLoss & vbCrLf
        DoEvents
        Server(Index).SendData "DISC" & strDisc & vbCrLf
        DoEvents
        Server(Index).SendData "TYPE" & strType(1) & vbCrLf
        DoEvents
        Server(Index).SendData "2TYPE" & strType(2) & vbCrLf
        DoEvents
        Server(Index).SendData "LVL" & strLvl & vbCrLf
        DoEvents
        
        Dim strKarma As String
        Dim strLastOn As String
        strLastOn = GetFromIni(curUser, "LASTON", nfile)
        strKarma = GetFromIni(curUser, "KARMA", nfile)
        If strKarma = "" Then strKarma = "0"
        If strLastOn <> strtime2 Then
            strKarma = CStr(CInt(strKarma) + 1)
        End If
        Call WriteIni(curUser, "KARMA", strKarma, nfile)
        
        Call WriteIni(curUser, "LASTON", strtime2, nfile)
        
        Server(Index).SendData "KARMA" & strKarma & vbCrLf
        DoEvents
        
        Server(Index).SendData "LOGINGOOD" & vbCrLf
        DoEvents
        

        
        
    For q = 1 To itotal
        Server(Index).SendData "CURITEM" & q & vbCrLf
        sTotal = GetFromIni("I" & q, "NAME", nSave)
        Server(Index).SendData "ITEMNAME" & sTotal & vbCrLf
        sTotal = GetFromIni("I" & q, "DESCRIPTION", nSave)
        Server(Index).SendData "ITEMDESC" & sTotal & vbCrLf
        sTotal = GetFromIni("I" & q, "DAMAGE", nSave)
        Server(Index).SendData "ITEMDMG" & sTotal & vbCrLf
        sTotal = GetFromIni("I" & q, "SPCDAMAGE", nSave)
        Server(Index).SendData "ITEMSPCDMG" & sTotal & vbCrLf
        sTotal = GetFromIni("I" & q, "SPCTYPE", nSave)
        Server(Index).SendData "ITEMSPCNAME" & sTotal & vbCrLf
        sTotal = GetFromIni("I" & q, "TYPE", nSave)
        Server(Index).SendData "ITEMTYPE" & sTotal & vbCrLf
        sTotal = GetFromIni("I" & q, "ADDMOD", nSave)
        Server(Index).SendData "ITEMADDMOD" & sTotal & vbCrLf
        sTotal = GetFromIni("I" & q, "MULTMOD", nSave)
        Server(Index).SendData "ITEMMULTMOD" & sTotal & vbCrLf
        sTotal = GetFromIni("I" & q, "SPCPERCENT", nSave)
        Server(Index).SendData "ITEMSPCPERCENT" & sTotal & vbCrLf
        sTotal = GetFromIni("I" & q, "COINS", nSave)
        Server(Index).SendData "ITEMCOINS" & sTotal & vbCrLf

        DoEvents
    Next 'i
    For p = 1 To 2
        nSave = App.Path & "\djinn.ini"
        If strType(p) = "E" Then
            sTotal = GetFromIni("GEN", "E", nSave)
            itotal = CInt(sTotal)
            curType = "E"
        End If
        If strType(p) = "F" Then
            sTotal = GetFromIni("GEN", "F", nSave)
            itotal = CInt(sTotal)
            curType = "F"
        End If
        If strType(p) = "N" Then
            sTotal = GetFromIni("GEN", "N", nSave)
            itotal = CInt(sTotal)
            curType = "N"
        End If
        If strType(p) = "W" Then
            sTotal = GetFromIni("GEN", "W", nSave)
            itotal = CInt(sTotal)
            curType = "W"
        End If
        If strType(p) = "H" Then
            sTotal = GetFromIni("GEN", "H", nSave)
            itotal = CInt(sTotal)
            curType = "H"
        End If
        If strType(p) = "D" Then
            sTotal = GetFromIni("GEN", "D", nSave)
            itotal = CInt(sTotal)
            curType = "D"
        End If
            
        For w = 1 To itotal
            'If CInt(strDjinn) >= w Then 'If I have enough Djinn in the first place
                Server(Index).SendData "CURDJINN" & (w + intLastDjinn) & vbCrLf
                DoEvents
                Server(Index).SendData "DJINNELEMENT" & strType(p) & vbCrLf
                DoEvents
                Server(Index).SendData "DJINNPLAYER" & p & vbCrLf
                DoEvents
                sTotal = GetFromIni(curType & w, "NAME", nSave)
                Server(Index).SendData "DJINNNAME" & sTotal & vbCrLf
                DoEvents
                sTotal = GetFromIni(curType & w, "DESCRIPTION", nSave)
                Server(Index).SendData "DJINNDESC" & sTotal & vbCrLf
                DoEvents
                sTotal = GetFromIni(curType & w, "TYPE", nSave)
                Server(Index).SendData "DJINNTYPE" & sTotal & vbCrLf
                DoEvents
                sTotal = GetFromIni(curType & w, "ADDMOD", nSave)
                Server(Index).SendData "DJINNADDMOD" & stotla & vbCrLf
                DoEvents
                sTotal = GetFromIni(curType & w, "DAMAGE", nSave)
                Server(Index).SendData "DJINNDMG" & sTotal & vbCrLf
                DoEvents
            'End If
        Next 'i
        intLastDjinn = itotal
            
        Server(Index).SendData "DJINNTOTAL" & strDjinn & vbCrLf
        DoEvents
            
        'Psynergy------------------------------------------------
    
        nSave = App.Path & "\psynergy.ini"
        
        sTotal = GetFromIni("GEN", curType, nSave)
        itotal = CInt(sTotal)
        
        Users(Index).Rating = CInt(strRating)
        iCurRank = CInt(strRating)
        
        If strType(p) = "E" Then
            sTotal = GetFromIni("GEN", "E", nSave)
            itotal = CInt(sTotal)
            curType = "E"
        End If
        If strType(p) = "F" Then
            sTotal = GetFromIni("GEN", "F", nSave)
            itotal = CInt(sTotal)
            curType = "F"
        End If
        If strType(p) = "D" Then
            sTotal = GetFromIni("GEN", "D", nSave)
            itotal = CInt(sTotal)
            curType = "D"
        End If
        If strType(p) = "W" Then
            sTotal = GetFromIni("GEN", "W", nSave)
            itotal = CInt(sTotal)
            curType = "W"
        End If
        If strType(p) = "H" Then
            sTotal = GetFromIni("GEN", "H", nSave)
            itotal = CInt(sTotal)
            curType = "H"
        End If
        If strType(p) = "D" Then
            sTotal = GetFromIni("GEN", "D", nSave)
            itotal = CInt(sTotal)
            curType = "D"
        End If
        
        For e = 1 To itotal
            strNewRank = GetFromIni(curType & e, "RATING", nSave)
            iNewRank = CInt(strNewRank)
            'If iCurRank >= iNewRank Then
                Server(Index).SendData "CURPSY" & (e + intLastPsy) & vbCrLf
                Server(Index).SendData "PSYELEMENT" & strType(p) & vbCrLf
                sTotal = GetFromIni(curType & e, "NAME", nSave)
                DoEvents
                Server(Index).SendData "PSYNAME" & sTotal & vbCrLf
                sTotal = GetFromIni(curType & e, "DAMAGE", nSave)
                DoEvents
                Server(Index).SendData "PSYDMG" & sTotal & vbCrLf
                sTotal = GetFromIni(curType & e, "DESC", nSave)
                DoEvents
                Server(Index).SendData "PSYDESC" & sTotal & vbCrLf
                sTotal = GetFromIni(curType & e, "TYPE", nSave)
                DoEvents
                Server(Index).SendData "PSYTYPE" & sTotal & vbCrLf
                sTotal = GetFromIni(curType & e, "PP", nSave)
                DoEvents
                Server(Index).SendData "PSYPP" & sTotal & vbCrLf
                DoEvents
                sTotal = GetFromIni(curType & e, "DJINN", nSave)
                Server(Index).SendData "PSYDJINN" & sTotal & vbCrLf
                DoEvents
                Server(Index).SendData "PSYPLAYER" & p & vbCrLf
                DoEvents
    
                Server(Index).SendData "DONE" & vbCrLf
            'End If
            DoEvents
        Next 'e
        intLastPsy = itotal
        
        nSave = App.Path & "\summons.ini"
        
        sTotal = GetFromIni("GEN", curType, nSave)
        itotal = CInt(sTotal)
        
        If strType(p) = "E" Then
            sTotal = GetFromIni("GEN", "E", nSave)
            itotal = CInt(sTotal)
            curType = "E"
        End If
        If strType(p) = "F" Then
            sTotal = GetFromIni("GEN", "F", nSave)
            itotal = CInt(sTotal)
            curType = "F"
        End If
        If strType(p) = "D" Then
            sTotal = GetFromIni("GEN", "D", nSave)
            itotal = CInt(sTotal)
            curType = "D"
        End If
        If strType(p) = "W" Then
            sTotal = GetFromIni("GEN", "W", nSave)
            itotal = CInt(sTotal)
            curType = "W"
        End If
        If strType(p) = "H" Then
            sTotal = GetFromIni("GEN", "H", nSave)
            itotal = CInt(sTotal)
            curType = "H"
        End If
        If strType(p) = "D" Then
            sTotal = GetFromIni("GEN", "D", nSave)
            itotal = CInt(sTotal)
            curType = "D"
        End If
    
        
        For r = 1 To itotal
            Server(Index).SendData "CURSUM" & (r + intLastSummon) & vbCrLf
            Server(Index).SendData "SUMELEMENT" & strType(p) & vbCrLf
            sTotal = GetFromIni(curType & r, "NAME", nSave)
            Server(Index).SendData "SUMNAME" & sTotal & vbCrLf
            sTotal = GetFromIni(curType & r, "DJINN", nSave)
            Server(Index).SendData "SUMDJINN" & sTotal & vbCrLf
            sTotal = GetFromIni(curType & r, "DESC", nSave)
            Server(Index).SendData "SUMDESC" & sTotal & vbCrLf
            DoEvents
            Server(Index).SendData "SUMCHAR" & p & vbCrLf
            DoEvents
        Next 'r
        intLastSummon = itotal
    
    Next 'p
    
    Dim cSave As String
    cSave = App.Path & "\customchar.ini"
    sTotal = GetFromIni("GEN", "TOTAL", cSave)
    itotal = CInt(sTotal)
    For p = 1 To itotal
        Server(Index).SendData "CUSTNUM" & CStr(p) & vbCrLf
        sTotal = GetFromIni(CStr(p), "NAME", cSave)
        Server(Index).SendData "CUSTNAME" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "DESCRIPTION", cSave)
        Server(Index).SendData "CUSTDESC" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "HP", cSave)
        Server(Index).SendData "CUSTHP" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "AP", cSave)
        Server(Index).SendData "CUSTAP" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "PP", cSave)
        Server(Index).SendData "CUSTPP" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "DEFENSE", cSave)
        Server(Index).SendData "CUSTDEFENSE" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "POWER", cSave)
        Server(Index).SendData "CUSTPOWER" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "RESIST", cSave)
        Server(Index).SendData "CUSTRES" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "STRENGTH", cSave)
        Server(Index).SendData "CUSTSTRENGTH" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "WEAKNESS", cSave)
        Server(Index).SendData "CUSTWEAKNESS" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "LUCK", cSave)
        Server(Index).SendData "CUSTLUCK" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "PICTURE", cSave)
        Server(Index).SendData "CUSTPICTURE" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "CLASS", cSave)
        Server(Index).SendData "CUSTTYPE" & sTotal & vbCrLf
        sTotal = GetFromIni(CStr(p), "USERS", cSave)
        Server(Index).SendData "CUSTUSER" & sTotal & vbCrLf
    Next 'p
    


    'Server(Index).SendData "GOOD" & vbCrLf

    timeWait.Enabled = False
    
    Else
    If ServerVersion = UserVersion Then
        Call WriteIni(strTime, "Data", "Bad Password Attempt (attempted password " & curPassword & ") on " & curUser & " at " & Server(Index).RemoteHostIP, xfile)
    Else
        Call WriteIni(strTime, "Data", "Bad version from " & curUser & " at " & Server(Index).RemoteHostIP, xfile)
    End If
        Server(Index).SendData "BAD" & vbCrLf
        DoEvents
'        Server(Index).Listen
        End If
        noclose = True
    loginWait = False
End If
If Left$(arrdata(i), 7) = "NEWUSER" Then
    NewUser = Mid(arrdata(i), 8, Len(arrdata(i)))
    noclose = True
End If
If Left$(arrdata(i), 6) = "NEWPIN" Then
    strNewUserPIN = Mid$(arrdata(i), 7, Len(arrdata(i)))
    Call WriteIni("PIN", strTime, NewUser & ": " & strNewUserPIN, xfile)
End If
If Left$(arrdata(i), 4) = "CHAR" Then
    strChar(1) = Mid(arrdata(i), 5, Len(arrdata(i)))
    noclose = True
    bCreateCustChar = False
End If
If Left$(arrdata(i), 5) = "2CHAR" Then
    strChar(2) = Mid(arrdata(i), 6, Len(arrdata(i)))
    noclose = True
    bCreateCustChar = False
End If
If Left$(arrdata(i), 8) = "CUSTCHAR" Then
    strChar(1) = Mid(arrdata(i), 7, Len(arrdata(i)))
    noclose = True
    bCreateCustChar = True
End If
If Left$(arrdata(i), 8) = "CUSTCHAR2" Then
    strChar(2) = Mid(arrdata(i), 7, Len(arrdata(i)))
    noclose = True
    bCreateCustChar = True
End If

If Left$(arrdata(i), 5) = "NEWPW" Then
    'comment out for ladder tournament
    
    NewPass = Mid(arrdata(i), 6, Len(arrdata(i)))
    
    bSwear = CheckSwear(NewUser)
    
    If NewPass = "celibi" Or NewPass = "alcatraz" Or bSwear = True Then
        bSave = App.Path & "\ban.ini"
        iBanMax = GetFromIni("GEN", "MAX", bSave)
        Call WriteIni("GEN", CStr(iBanMax + 1), Server(Index).RemoteHostIP, bSave)
        Call WriteIni("GEN", "MAX", CStr(iBanMax + 1), bSave)
        Server(Index).Close
        Call WriteIni(CStr(iBanMax + 1), "New User", strTime, App.Path & "\zidane.ini")
    End If
    
    intPinBanMax = CInt(GetFromIni("GEN", "TOTAL", App.Path & "\pinban.ini"))
    strCurPin = strNewUserPIN
    For q = 1 To intPinBanMax
        strTempPin = GetFromIni("GEN", CStr(q), App.Path & "\pinban.ini")
        If InStr(1, strCurPin, strTempPin, vbTextCompare) Then
            Server(Index).SendData "BADPIN" & vbCrLf
            Server(Index).Close
            Exit Sub
        End If
    Next 'q
    
    Dim getMax As Integer
    Dim IsItNew As String
    
    IsItNew = ""
    'getMax = CInt(GetFromIni("GEN", "TOTAL", nfile))
    'For q = 1 To getMax
    '    strusercheck = GetFromIni("USERNUM", CStr(q), nfile)
    '    If NewUser = strusercheck Then IsItNew = "NOT"
    '    DoEvents
    'Next 'q

    strusercheck = GetFromIni(NewUser, "PASSWORD", nfile)
    If strusercheck <> "" Then
        IsItNew = "NOT"
    Else
        IsItNew = ""
    End If
    
    Dim uSave As String
    uSave = App.Path & "\nouse.ini"
    Dim intTotalNoUse As Long
    strTotalNoUse = GetFromIni("GEN", "TOTAL", uSave)
    If strTotalNoUse = "" Then strTotalNoUse = "0"
    intTotalNoUse = CInt(strTotalNoUse)
    Dim strNoUse As String
    Dim bCheckNoUse As Long
    bCheckNoUse = 0
    For q = 1 To intTotalNoUse
        strNoUse = GetFromIni(CStr(q), "NAME", uSave)
        If InStr(NewUser, strNoUse) <> 0 Then
            bCheckNoUse = 999
        End If
    Next 'q
    
    
    If IsItNew = "" And bCheckNoUse = 0 Then
    sTotal = GetFromIni("GEN", "TOTAL", nfile)
    itotal = CInt(sTotal)
    itotal = itotal + 1
    sTotal = CStr(itotal)
    Call WriteIni("GEN", "TOTAL", sTotal, nfile)
    Call WriteIni("GEN", "NEWEST", NewUser, nfile)
    Call WriteIni("USERNUM", sTotal, NewUser, nfile)
    
    Call WriteIni(NewUser, "PASSWORD", NewPass, nfile)
    Call WriteIni(NewUser, "LEVEL", "1", nfile)
    Call WriteIni(NewUser, "DJINNNUM", "1", nfile)
    Call WriteIni(NewUser, "COINS", "100", nfile)
    Call WriteIni(NewUser, "RATING", "1000", nfile)
    Call WriteIni(NewUser, "CHAR", strChar(1), nfile)
    Call WriteIni(NewUser, "WINS", "0", nfile)
    Call WriteIni(NewUser, "DISC", "0", nfile)
    Call WriteIni(NewUser, "LOSS", "0", nfile)
    Call WriteIni(NewUser, "ITEM", "1", nfile)
    Call WriteIni(NewUser, "DPOINTS", "0", nfile)
    Call WriteIni(NewUser, "NAME", NewUser, nfile)
        
    If strChar(1) = "Isaac" Or strChar(1) = "Guard" Or strChar(1) = "Gladiator" Then
    Call WriteIni(NewUser, "TYPE", "E", nfile)
    End If
    If strChar(1) = "Kenny" Or strChar(1) = "Jenna" Or strChar(1) = "Garret" Or strChar(1) = "Saturos" Or strChar(1) = "Menardi" Then
    Call WriteIni(NewUser, "TYPE", "F", nfile)
    End If
    If strChar(1) = "Ivan" Or strChar(1) = "Sheba" Or strChar(1) = "Cloud" Then
    Call WriteIni(NewUser, "TYPE", "N", nfile)
    End If
    If strChar(1) = "Purple Piers" Or strChar(1) = "Piers" Or strChar(1) = "Mia" Or strChar(1) = "Alex" Or strChar(1) = "Caption Contest Character" Then
    Call WriteIni(NewUser, "TYPE", "W", nfile)
    End If
    If strChar(1) = "Felix" Or strChar(1) = "The Wise One" Then
    Call WriteIni(NewUser, "TYPE", "H", nfile)
    End If
    If strChar(1) = "Kraden" Or strChar(1) = "KOS" Or strChar(1) = "Karst" Or strChar(1) = "Agiato" Then
    Call WriteIni(NewUser, "TYPE", "D", nfile)
    End If
    If bCreateCustChar = True Then
        Dim strCustCharType As String
        strCustCharType = GetFromIni(strChar(1), "CLASS", App.Path & "\customchar.ini")
        Call WriteIni(NewUser, "TYPE", strCustCharType, nfile)
        strCustCharType = GetFromIni(strChar(1), "NAME", App.Path & "\customchar.ini")
        Call WriteIni(NewUser, "CHAR", strCustCharType, nfile)
    End If
    
    
    Server(Index).SendData "GUSR" & vbCrLf


    DoEvents
    
    noclose = True
    closeme.Enabled = True
    Call WriteIni(strTime, "Data", "Created new user " & NewUser & " at " & Server(Index).RemoteHostIP, xfile)
    Else
    noclose = True
    Call WriteIni(strTime, "Data", "Bad password attempt on new user " & NewUser & " at " & Server(Index).RemoteHostIP, xfile)
    Server(Index).SendData "BUSR" & vbCrLf
    Debug.Print NewPass
    Debug.Print NewUser
'    Timeout (2)

'    Timeout (2)
'    Server(Index).Listen
    End If
    'closeme.Enabled = True
End If

If Left$(arrdata(i), 7) = "WINUSER" Then
'    noclose = True
    nfile = App.Path & "\user.ini"
    curUser = Mid(arrdata(i), 8, Len(arrdata(i)))
    
    
'    ServerNum(Index) = CInt(curUser)
    
    
'    DoEvents

'    Dim strTestName As String
'    strTestName = GetFromIni(curuser, "NAME", nfile)
'    If curUser <> strTestName And curUser <> "" Then
'        ServerNum(Index) = FindUser(curUser)
'    End If
    




    Call WriteIni(strTime, "USERNUM", curUser, App.Path & "\gamelog.ini")
    

End If

    If Left$(arrdata(i), 6) = "RATING" Then
        curUser = Mid(arrdata(i), 7, Len(arrdata(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)
        strCurRating = GetFromIni(curUser, "RATING", nfile)
        iCurRating = CInt(strCurRating)
        Dim snewRating As String
        Dim inewRating As Integer
        snewRating = CStr(strsplit(1))
        inewRating = CInt(snewRating)
        
        Call WriteIni(strGameTime, "Gained Rating", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        
        snewRating = CStr(inewRating + iCurRating)
        iCurRating = inewRating + iCurRating
        iCurRating = iCurRating - 1000
        Dim intCurLvl As Long
        Dim intCurDjinn As Long
        intCurLvl = GetLevel(CStr(iCurRating))
        intCurDjinn = GetDjinn(CStr(iCurRating))
        Call WriteIni(curUser, "RATING", snewRating, nfile)
        Call WriteIni(curUser, "LEVEL", CStr(intCurLvl), nfile)
        Call WriteIni(curUser, "DJINNNUM", CStr(intCurDjinn), nfile)
    End If
    
    If Left$(arrdata(i), 5) = "COINS" Then
        curUser = Mid(arrdata(i), 6, Len(arrdata(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)
        
        Dim snewCoins As String
        Dim inewCoins As Integer
        strCurCoins = GetFromIni(curUser, "COINS", nfile)
        iCurCoins = CInt(strCurCoins)
        snewCoins = CStr(strsplit(1))
        
        
        Call WriteIni(strGameTime, "Gained Coins", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        
        inewCoins = CInt(snewCoins)
        snewCoins = CStr(inewCoins + iCurCoins)
        Call WriteIni(curUser, "COINS", snewCoins, nfile)
        
    End If
    'If Left$(arrdata(i), 11) = "DJINNGAINED" Then
    '    curUser = Mid(arrdata(i), 12, Len(arrdata(i)))
    '    strsplit = Split(curUser, "@", -1, vbTextCompare)
    '    curUser = strsplit(0)

     '   Dim sDjinnGained As String
     '   Dim iDjinnGained As Integer
     '   sCurDjinn = GetFromIni(curUser, "DJINNNUM", nfile)
     '   iCurDjinn = CInt(sCurDjinn)
     '   sDjinnGained = CStr(strsplit(1))
     '   Call WriteIni(strTime, "DJINN-GAINED", curUser & sDjinnGained, App.Path & "\gamelog.ini")
     '   iDjinnGained = CInt(sDjinnGained)
     '   sDjinnGained = CStr(iDjinnGained + iCurDjinn)
     '   Call WriteIni(curUser, "DJINNNUM", sDjinnGained, nfile)
        
    'End If
    'If Left$(arrdata(i), 3) = "LVL" Then
    '    curUser = Mid(arrdata(i), 4, Len(arrdata(i)))
    '    strsplit = Split(curUser, "@", -1, vbTextCompare)
    '    curUser = strsplit(0)
    '    Dim snewLVL As String
    '    Dim inewLVL As Integer
      '  strCurLvl = GetFromIni(curUser, "LEVEL", nfile)
      '  iCurLvl = CInt(strCurLvl)
      '  snewLVL = CStr(strsplit(1))
      '  Call WriteIni(strTime, "LVL-GAINED", curUser & snewLVL, App.Path & "\gamelog.ini")
      '  inewLVL = CInt(snewLVL)
       ' snewLVL = CStr(inewLVL + iCurLvl)
      '  Server(Index).SendData "STAT" & vbCrLf
     '   Call WriteIni(curUser, "LEVEL", snewLVL, nfile)

    'End If
    If Left$(arrdata(i), 4) = "SWIN" Then
        curUser = Mid(arrdata(i), 5, Len(arrdata(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)

        Call WriteIni(strGameTime, "Won Against", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        Dim CurWins As String
        Dim iWins As Integer
        CurWins = GetFromIni(curUser, "WINS", nfile)
        iWins = CInt(CurWins)
        CurWins = CStr(iWins + 1)
        Call WriteIni(curUser, "WINS", CurWins, nfile)
    End If
    If Left$(arrdata(i), 4) = "LOSE" Then
        curUser = Mid(arrdata(i), 5, Len(arrdata(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)

        Call WriteIni(strGameTime, "Lost Against", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        Dim CurLose As String
        Dim iLose As Integer
        CurLose = GetFromIni(curUser, "LOSS", nfile)
        iLose = CInt(CurLose)
        CurLose = CStr(iLose + 1)
        Call WriteIni(curUser, "LOSS", CurLose, nfile)
    End If
    
If Left$(arrdata(i), 9) = "STATSLOSS" Then
    nfile = App.Path & "\user.ini"
    'Dim curUser As String
    curUser = Mid(arrdata(i), 10, Len(arrdata(i)))
    
    curMax = CInt(GetFromIni("GEN", "TOTAL", nfile))
    For q = 0 To curMax
        chkuser = GetFromIni(CStr(q), "NAME", nfile)
        If chkuser = curUser Then
        curUser = CStr(q)
        End If
        'DoEvents
    Next 'i
    
    strCurRating = GetFromIni(curUser, "RATING", nfile)
    iCurRating = CInt(strCurRating)
    strCurLvl = GetFromIni(curUser, "LEVEL", nfile)
    iCurLvl = CInt(strCurLvl)
    strCurCoins = GetFromIni(curUser, "COINS", nfile)
    iCurCoins = CInt(strCurCoins)
End If

If Left$(arrdata(i), 10) = "LOSSRATING" Then
        curUser = Mid(arrdata(i), 11, Len(arrdata(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)
        strCurRating = GetFromIni(curUser, "RATING", nfile)
        iCurRating = CInt(strCurRating)

        snewRating = CStr(strsplit(1))
        inewRating = CInt(snewRating)
        
        Call WriteIni(strGameTime, "Lost Rating", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        
        snewRating = CStr(iCurRating - inewRating)
        iCurRating = iCurRating - inewRating
        iCurRating = iCurRating - 1000
        If iCurRating < 0 Then iCurRating = 0

        intCurLvl = GetLevel(CStr(iCurRating))
        intCurDjinn = GetDjinn(CStr(iCurRating))
        Call WriteIni(curUser, "RATING", snewRating, nfile)
        Call WriteIni(curUser, "LEVEL", CStr(intCurLvl), nfile)
        Call WriteIni(curUser, "DJINNNUM", CStr(intCurDjinn), nfile)
    
End If

If Left$(arrdata(i), 9) = "LOSSCOINS" Then
        curUser = Mid(arrdata(i), 10, Len(arrdata(i)))
        strsplit = Split(curUser, "@", -1, vbTextCompare)
        curUser = strsplit(0)

        strCurCoins = GetFromIni(curUser, "COINS", nfile)
        iCurCoins = CInt(strCurCoins)
        snewCoins = CStr(strsplit(1))
        
        
        Call WriteIni(strGameTime, "Lost Coins", CStr(strsplit(1)), strSambarPath & "\" & curUser & ".ini")
        
        
        inewCoins = CInt(snewCoins)
        snewCoins = CStr(iCurCoins - inewCoins)
        Call WriteIni(curUser, "COINS", snewCoins, nfile)
End If


    
If Left$(arrdata(i), 10) = "SINGLENAME" Then

    snewName = Mid(arrdata(i), 11, Len(arrdata(i)))
End If
If Left$(arrdata(i), 11) = "SINGLECOINS" Then
    Dim ssingCoins As String
    ssingCoins = Mid(arrdata(i), 12, Len(arrdata(i)))
    Dim sSingCurCoins As String
    sSingCurCoins = GetFromIni(snewName, "COINS", App.Path & "\user.ini")
    Call WriteIni(strTime, "COINSIMPORTED", ssingCoins, App.Path & "\data.ini")
    ssingCoins = CStr(CInt(ssingCoins) + CInt(sSingCurCoins))
    Call WriteIni(snewName, "COINS", ssingCoins, App.Path & "\user.ini")
    Server(Index).SendData "SINGLECOINS" & vbCrLf
End If
    


If Left$(arrdata(i), 11) = "NEWITEMUSER" Then
    Dim strUserGuy As String
    strUserGuy = Mid(arrdata(i), 12, Len(arrdata(i)))
End If

If Left$(arrdata(i), 11) = "NEWITEMCHAR" Then
    Dim strItemPlayer As String
    strItemPlayer = Mid(arrdata(i), 12, Len(arrdata(i)))
End If

If Left$(arrdata(i), 12) = "NEWITEMCOINS" Then
    Dim strNewCoins As String
    strNewCoins = Mid(arrdata(i), 13, Len(arrdata(i)))
    Call WriteIni(strUserGuy, "COINS", strNewCoins, nfile)
End If

If Left$(arrdata(i), 11) = "NEWITEMNAME" Then
    Dim strNewItemName As String
    strNewItemName = Mid(arrdata(i), 12, Len(arrdata(i)))
    Dim strItemTotal As String
    Dim iItemtotal As Integer
    strItemTotal = GetFromIni("GEN", "TOTAL", App.Path & "\items.ini")
    iItemtotal = CInt(strItemTotal)
    Dim strCheckItemName As String
    For q = 1 To iItemtotal
        strCheckItemName = GetFromIni("I" & q, "NAME", App.Path & "\items.ini")
    If strNewItemName = strCheckItemName Then
        If strItemPlayer = "1" Then
            Call WriteIni(strUserGuy, "ITEM", CStr(q), nfile)
        Else
            Call WriteIni(strUserGuy, "ITEM2", CStr(q), nfile)
        End If
    End If
    DoEvents
    Next 'q
    Server(Index).SendData "ITEMCONFIRM" & vbCrLf
End If
If Left$(arrdata(i), 9) = "ERRORDESC" Then
    Dim eSave As String
    Dim strErrorDesc As String
    eSave = App.Path & "\errorlog.ini"
    strErrorDesc = Mid(arrdata(i), 10, Len(arrdata(i)))
    Call WriteIni(strTime, "ERROR", strErrorDesc, eSave)
End If
If Left$(arrdata(i), 8) = "ERRORNUM" Then
    eSave = App.Path & "\errorlog.ini"
    strErrorDesc = Mid(arrdata(i), 9, Len(arrdata(i)))
    Call WriteIni(strTime, "NUM", strErrorDesc, eSave)
End If
If Left$(arrdata(i), 11) = "ERRORSOURCE" Then
    eSave = App.Path & "\errorlog.ini"
    strErrorDesc = Mid(arrdata(i), 12, Len(arrdata(i)))
    Call WriteIni(strTime, "SOURCE", strErrorDesc, eSave)
End If
If Left$(arrdata(i), 12) = "HIGHPLAYERDS" Then
    strDjinnSavePlayer = Mid$(arrdata(i), 13, Len(arrdata(i)))
End If
If Left$(arrdata(i), 11) = "HIGHSCOREDS" Then
    intDjinnSaveHighScore = CInt(Mid(arrdata(i), 12, Len(arrdata(i))))
    For q = 1 To 20
        If Users(q).Enabled = True Then
            Server(q).SendData "HIGHSCOREDS" & CStr(intDjinnSaveHighScore) & vbCrLf
            Server(q).SendData "HIGHPLAYERDS" & strDjinnSavePlayer * vbCrLf
        End If
    Next ' q
    Call WriteIni("GEN", "SCORE", CStr(intDjinnSaveHighScore), App.Path & "\motd.ini")
    Call WriteIni("GEN", "PLAYER", strDjinnSavePlayer, App.Path & "\motd.ini")
End If

If Left$(arrdata(i), 7) = "RQSTLST" Then        'if user requested file list
        strTemp = ""
        strFileList = ""
        filChar.Path = App.Path & "\files\"  'get the first filename
        If filChar.List(0) = "" Then                    'if there are no files then
            Server(Index).SendData "LSTnul"    'send the nofiles message (nul is a virual file that exists in every directory)
        Else                                    'else if there are files then start collecting a list of all files
            Call FileListLoop(0, Index, True)
        End If
End If
If Left$(arrdata(i), 6) = "RQSTFL" Then        'if user requested file
    Dim strFile As String
    Dim strNewFile As String
    Dim strEncodedFile As String
    strFile = Mid$(arrdata(i), 7, Len(arrdata(i)))
    strNewFile = App.Path & "\files\" & strFile
    Dim b64 As New base64
    strEncodedFile = b64.EncodeFromFile(strNewFile)
    Server(Index).SendData "FILE" & strFile & "@" & strEncodedFile & "!" 'send the file
End If

    If Left$(arrdata(i), 12) = "GETDJINNCHAR" Then
        Dim strGetDjinnChar As String
        strGetDjinnChar = Mid(strData(i), 13, Len(strData(i)))
    End If
    If Left$(arrdata(i), 11) = "GETDJINNNUM" Then
        Dim intGetDjinnNum As Long
        intGetDjinnNum = CLng(Mid(strData(i), 12, Len(strData(i))))
        curType = strGetDjinnChar
        nSave = App.Path & "\djinn.ini"
        
        Server(Index).SendData "CURDJINN" & w & vbCrLf
        DoEvents
        sTotal = GetFromIni(curType & w, "NAME", nSave)
        Server(Index).SendData "DJINNNAME" & sTotal & vbCrLf
        DoEvents
        sTotal = GetFromIni(curType & w, "DESCRIPTION", nSave)
        Server(Index).SendData "DJINNDESC" & sTotal & vbCrLf
        DoEvents
        sTotal = GetFromIni(curType & w, "TYPE", nSave)
        Server(Index).SendData "DJINNTYPE" & sTotal & vbCrLf
        DoEvents
        sTotal = GetFromIni(curType & w, "ADDMOD", nSave)
        Server(Index).SendData "DJINNADDMOD" & stotla & vbCrLf
        DoEvents
        sTotal = GetFromIni(curType & w, "DAMAGE", nSave)
        Server(Index).SendData "DJINNDMG" & sTotal & vbCrLf
        DoEvents
    End If



'DoEvents

Next 'i

    If arrdatao <> "CHATSPAM" & vbCrLf Then
        lstStatus.AddItem strTime & " " & arrdatao
    End If
    
Exit Sub

err:
On Error Resume Next
Debug.Print err.Description
'Exit Sub


End Sub

Private Sub Server_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
Call WriteIni(strTime, "Errors", Description, App.Path & "\data.ini")
Server(Index).Close
'Server(Index).Listen
End Sub

Private Sub timeclose_Timer()
On Error Resume Next
lblstatus.Caption = Server(Index).State
End Sub

Private Sub PopulateUser(ByVal q As Integer)
On Error Resume Next
    Server(q).SendData "CHATTXT" & "Message of the Day:" & txtMOTD.Text & vbCrLf
    
    
    For i = 1 To 20
    If UserName(i) <> "" And i <> q Then
    Server(q).SendData "CHATNUM" & i & vbCrLf
    Server(q).SendData "CHATNAME" & UserName(i) & vbCrLf
    Server(q).SendData "CHATRATING" & UserRating(i) & vbCrLf
    Server(q).SendData "CHATIP" & Server(i).RemoteHostIP & vbCrLf
'    Server(q).SendData "ISAACPIC" & Users(i).Pic & vbCrLf
    
    DoEvents
    End If
    
    If UserName(q) <> "" Then
    Server(i).SendData "CHATNUM" & iNewServer & vbCrLf
    Server(i).SendData "CHATNAME" & UserName(q) & vbCrLf
    Server(i).SendData "CHATRATING" & UserRating(q) & vbCrLf
    Server(i).SendData "CHATIP" & Server(q).RemoteHostIP & vbCrLf
    DoEvents
    End If
    DoEvents
    Next 'i
End Sub
Sub DestroyUser(ByVal q As String)
On Error Resume Next
UserName(q) = ""
UserRating(q) = ""
For i = 1 To 20
Server(i).SendData "CHATKILL" & q & vbCrLf
DoEvents
Next 'i
End Sub
Sub SendChat(ByVal strChatTxt As String)
On Error Resume Next

For q = 1 To 20
If Users(q).Enabled = True Then
Chat(q).SendData "CHATTXT" & strChatTxt & vbCrLf
DoEvents
End If
'DoEvents
Next 'q

End Sub

Private Sub timeFilePing_Timer()
On Error GoTo err
For i = 1 To 7
    If FileTransfer(i).State <> sckClosed Then
        FileTransfer(i).SendData "PING"
        DoEvents
    End If
Next 'i
Exit Sub
err:
FileTransfer(i).Close
End Sub

Private Sub timeMulti_Timer()
For i = 1 To 20
    If Users(i).Enabled = True Then
        Call AddMUser(i, False)
    End If
    DoEvents
Next 'i
End Sub

Private Sub timePing_Timer()
On Error GoTo err

filChar.Path = ""
filChar.Path = App.Path & "\files\"

Dim intSockets As Integer
intSockets = 0
For i = 1 To 20
    If Users(i).Enabled = True Then
        intSockets = intSockets + 1
        Chat(i).SendData "P" & vbCrLf
    End If
    DoEvents
Next 'i
lblgen(38).Caption = intSockets

Call SendUsers


Exit Sub
err:
If Users(i).Name <> "" Then
Call SendDataToAll("AWAYTXT" & Users(i).Name & " has disconnected.")
End If
Users(i).Enabled = False
Chat(i).Close
Game(i).Enabled = False
Call KillChar(CStr(i))
Resume Next

End Sub

Private Sub timeServerPing_Timer()

'On Error GoTo err
For i = 0 To 20
    If Server(i).State = sckConnected Then
        Server(i).SendData "P" & vbCrLf
    End If
    DoEvents
Next 'i
err:
If i <= 20 Then
Server(i).Close
DoEvents
    If i = 0 Then
        If Server(i).State = sckClosed Then
        Server(i).Listen
        End If
    Else
        Server(i).Close
    End If
'Resume Next
Else
'Resume Next
End If

End Sub

Private Sub timeUpdate_Timer()
On Error Resume Next
'timeUpdate.Enabled = False
ladderRefresh = ladderRefresh + 1

'lstRank.Clear
'txtHTML.Text = ""

If ladderRefresh Mod 6 = 0 Then
    upTime = upTime + 1 'Server uptime
End If

If ladderRefresh >= 1800 Then
    Call UpdateLadder
End If

Debug.Print Chat(0).State
If Chat(0).State <> 2 Then
    Chat(0).Close
    DoEvents
    Chat(0).Listen
End If

If ladderRefresh Mod 90 = 0 And cmdCloseServer.Caption <> "Re-Open Server" Then
    strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
        txtHTML.Text = ""
        Dim intTotal As Integer
        For i = 1 To 20
            If Users(i).Enabled = True Then
                intTotal = intTotal + 1
            End If
            DoEvents
        Next 'i
        strtotal = CStr(intTotal)
        struptime = CStr(upTime)
        strLastUser = GetFromIni("GEN", "NEWEST", App.Path & "\user.ini")
        Dim strScrambler As String
        If chkScrambler.Value = 1 Then
            strScrambler = "On"
        Else
            strScrambler = "Off"
        End If
        txtHTML.Text = "This webpage is updated about every 15 minutes if the server is up.  If it hasn't been updated within the last few minutes, then the server is probably down for maitence.  Please don't post issues about the server being down on the Bug Report forum.<br><br>As of <b>" & strTime & "</b> the server is running <b>Version " & txtVer.Text & "</b> and has been running for <b>" & CStr(CInt(struptime) * 15) & "</b> minutes.<br><br>There are currently <b>" & strtotal & "</b> users online.<br>The last user to register was <b>" & strLastUser & "</b><br>Scrambler is currently: <b>" & strScrambler & "</b><br><br>Message of the Day: <b>" & txtMOTD.Text & "</b>"
 FNO = FreeFile
     On Error Resume Next
     err.Clear
     Open (strSambarPath & "\status.html") For Output As #FNO
      If err.Number <> 0 Then
        'MsgBox "an error has occured"
       Else
        Print #FNO, (txtHTML.Text)
      End If
     Close #FNO
     'On Error GoTo 0
End If

End Sub

Private Sub timeWait_Timer()
On Error Resume Next
loginWait = False
timeWait.Enabled = False
End Sub

Private Sub txtChat_Change()
On Error Resume Next
If Len(txtChat.Text) >= 5000 Then
    txtChat.Text = Right$(txtChat.Text, 5000)
End If
Call AutoScroll(txtChat)
End Sub

Private Sub AutoScroll(Box As TextBox)
Box.SelStart = Len(Box.Text) - 1
End Sub

Private Sub txtChat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    txtChat.Text = ""
End If
End Sub

Private Sub txtMsg_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    Call cmdSend_Click
End If

End Sub

Private Sub txtVer_Change()
On Error Resume Next
Call WriteIni("MOTD", "VER", txtVer.Text, App.Path & "\motd.ini")
ServerVersion = txtVer.Text
End Sub
Sub AdminChat(ByVal strChatTxt As String)
On Error Resume Next
For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "ADMINTXT" & strChatTxt & vbCrLf
        DoEvents
    End If
Next 'q
End Sub
Sub ChatDisc()
On Error Resume Next
For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "CHATKILL" & strDiscUser & vbCrLf
        DoEvents
    End If
Next ' q
End Sub
Sub ChatSpam()
On Error Resume Next
For z = 1 To 20
    If Users(z).Enabled = True Then
        Chat(z).SendData "CHATNUM" & z & vbCrLf
        Chat(z).SendData "CHATNAME" & Users(z).Name & vbCrLf
        Chat(z).SendData "CHATRATING" & Users(z).Rating & vbCrLf
        Chat(z).SendData "CHATIP" & Users(z).IP & vbCrLf
        DoEvents
    End If
Next 'z
End Sub
Sub ChatMsg(ByVal strMsg As String, ByVal strUser As String)
On Error Resume Next
For q = 1 To 20
    If Users(q).Name = strUser Then
    
        Chat(q).SendData "CHATMSG" & "[You Recieved a Private Message]: " & strMsg & vbCrLf
        DoEvents
    End If
Next 'q
End Sub
Sub SendUsers()
On Error Resume Next

For i = 1 To 20
If Users(i).Enabled = True Then
    Chat(i).SendData "CHATSTART" & vbCrLf
    For q = 1 To 20
        If Users(q).Enabled = True Then
            Chat(i).SendData "CHATNAME" & Users(q).Name & vbCrLf
            DoEvents
        End If
    Next 'q
    Chat(i).SendData "CHATSTOP" & vbCrLf
End If
Next 'i

End Sub

Sub AddMUser(ByVal intCur As Integer, FirstLoad As Boolean)
On Error Resume Next

For i = 1 To 20
    If Users(i).Enabled = True Then
        Chat(i).SendData "ISAACCURNUM" & intCur & vbCrLf
        DoEvents
        If FirstLoad = True Then
            Chat(intCur).SendData "ISAACPIC" & Users(i).Pic & vbCrLf
            DoEvents
            Chat(i).SendData "ISAACPIC" & Users(intCur).Pic & vbCrLf
            DoEvents
        End If
        Chat(i).SendData "ISAACSCREEN" & Users(intCur).Screen & vbCrLf
        DoEvents
        Chat(i).SendData "PICISAAC" & Users(intCur).Pic & vbCrLf
        DoEvents
        Chat(i).SendData "MOVEISAACX" & Users(intCur).Left & vbCrLf
        DoEvents
        Chat(i).SendData "MOVEISAACY" & Users(intCur).Top & vbCrLf
        DoEvents
        If FirstLoad = True Then
        Chat(i).SendData "AVATAR" & CStr(Index) & "@" & Users(intCur).Avatar & vbCrLf
        DoEvents
        End If
    End If
    If i = intCur Then
        For q = 1 To 20
            Chat(i).SendData "USERNAME" & CStr(q) & "@" & Users(q).Name & vbCrLf
            DoEvents
            Chat(i).SendData "AVATAR" & CStr(q) & "@" & Users(q).Avatar & vbCrLf
            DoEvents
        Next 'q
    End If
Next 'i

End Sub
Sub KillChar(Player As Integer)
On Error Resume Next
For i = 1 To 20
    If Users(i).Enabled = True Then
        Chat(i).SendData "ISAACKILL" & Player & vbCrLf
        DoEvents
    End If
Next 'i
End Sub
Sub GetGameList(Client As Integer)
On Error Resume Next
For i = 1 To 20
    If Game(i).Enabled = True Then
        Chat(Client).SendData "GAMEADD" & Game(i).Name & vbCrLf
    End If
    DoEvents
Next 'i
End Sub
Sub JoinGame(ByVal GameName As String, ByVal Index As Integer)
On Error Resume Next
For i = 1 To 20
    If GameName = Game(i).Name Then
        Chat(Index).SendData "JOINIP" & Game(i).IP
    End If
    DoEvents
Next 'i
End Sub
Private Function FindUser(UserName As String) As Integer
On Error Resume Next
    Dim intMax As Integer
    Dim nSave As String
    
    nSave = App.Path & "\user.ini"
    intMax = CInt(GetFromIni("GEN", "TOTAL", nSave))
    For i = 0 To intMax
        Dim strtempName As String
        strtempName = GetFromIni(CStr(i), "NAME", nSave)
        If strtempName = UserName Then
            FindUser = i
            Exit Function
        End If
        DoEvents
    Next 'i
    
End Function
Sub InGameChat(ByVal Index As Integer, ByVal strChat As String)
On Error Resume Next
    For i = 1 To 20
        If Users(i).Enabled = True Then
            Chat(i).SendData "INGAMECHATNUM" & CStr(Index) & vbCrLf
            Chat(i).SendData "INGAMECHATTXT" & strChat & vbCrLf
        End If
        DoEvents
    Next 'i
End Sub
Sub SendMeChat(ByVal strChatTxt As String)
On Error Resume Next

For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "METXT" & strChatTxt & vbCrLf
        DoEvents
    End If
    DoEvents
Next 'q

End Sub
Sub SendSndChat(ByVal strChatTxt As String)
On Error Resume Next

For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "CHATSND" & strChatTxt & vbCrLf
        DoEvents
    End If
    DoEvents
Next 'q

End Sub
Sub SendInChat(ByVal strNumber As String, ByVal Index As Integer)
On Error Resume Next
        For q = 1 To 20
            If Users(q).Enabled = True Then
                Chat(q).SendData "INICONNUM" & Index & vbCrLf
                DoEvents
                Chat(q).SendData "INICONTYPE" & strNumber & vbCrLf
                DoEvents
            End If
            DoEvents
        Next ' q
End Sub
Sub ModChat(ByVal strChatTxt As String)
On Error Resume Next
For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "MODTXT" & strChatTxt & vbCrLf
        DoEvents
    End If
Next 'q
End Sub
Sub FileListLoop(intCurFile As Long, Index As Integer, bServer As Boolean)
On Error Resume Next
        strTemp = filChar.List(intCurFile)   'get the next filename
        If strTemp <> "" Then                'if there was annother file then
            If strFileList <> "" Then
                strFileList = strFileList & "," & strTemp 'add it to the list and
            Else
                strFileList = strTemp
            End If
            Call FileListLoop(intCurFile + 1, Index, bServer)                    'continue the loop
            Exit Sub
        Else                                  'otherwise
            If bServer = False Then
                FileTransfer(Index).SendData "LST" & strFileList     'send the filelist to the client
            Else
                Server(Index).SendData "LST" & strFileList     'send the filelist to the client
            End If
        End If
End Sub
Sub SendList(ByVal Index As Integer, ByVal bServer As Boolean)
On Error Resume Next
strTemp = ""
strFileList = ""
filChar.Path = App.Path & "\files\"  'get the first filename
If filChar.List(0) = "" Then                    'if there are no files then
    FileTransfer(Index).SendData "LSTnul"     'send the nofiles message (nul is a virual file that exists in every directory)
Else                                    'else if there are files then start collecting a list of all files
    Call FileListLoop(0, Index, bServer)
End If

End Sub
Sub TalkChat(ByVal strChatTxt As String)
On Error Resume Next

For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "TALKTXT" & strChatTxt & vbCrLf
        DoEvents
    End If
    DoEvents
Next 'q

End Sub
Sub GameChat(ByVal strChatTxt As String, bMe As Boolean)
On Error Resume Next

For q = 1 To 20
    If Users(q).Enabled = True Then
        If bMe = False Then
        Chat(q).SendData "GAMETXT" & strChatTxt & vbCrLf
        Else
        Chat(q).SendData "GAMEMETXT" & strChatTxt & vbCrLf
        End If
        DoEvents
    End If
    'DoEvents
Next 'q

End Sub
Private Function CheckSwear(ByVal strText As String) As Boolean
On Error Resume Next
Dim strCurse As String
Dim bSwear As Boolean
Dim strReturn As Long
bSwear = False
For i = 1 To 11
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
        strCurse = "server"
    End If
    Dim strTxt As String
    strTxt = LCase(strText)
    strReturn = InStr(1, strTxt, strCurse, vbTextCompare)
    If strReturn > 0 Then
        bSwear = True
    End If
Next 'i
If bSwear = True Then CheckSwear = True
If bSwear = False Then CheckSwear = False
End Function

Function GetLevel(strR As String) As Long
On Error Resume Next
Dim intTemp As Long
Dim intTotal As Long

intTemp = CLng(strR)

If intTemp < 0 Then
    GetLevel = 1
    Exit Function
End If

For i = 1 To 100
    intTotal = intTotal + (25 + i ^ 1.5)
    If intTotal > intTemp Then
        GetLevel = i
        Exit Function
    End If
Next 'i

'Formula for gaining a level
'Rating To Next Level = 25 + iLvl ^ 1.5


End Function
Function GetDjinn(strR As String) As Long
On Error Resume Next
Dim intTemp As Long
Dim intTotal As Long

intTemp = CLng(strR)


For i = 1 To 100
    intTotal = intTotal + (25 + i ^ 1.5)
    If intTotal > intTemp Then
        GetDjinn = i / 2
        Exit Function
    End If
Next 'i

End Function
Private Sub DownloadNextFile(Index As Long)
On Error Resume Next
If intCurFile < UBound(FileNameList) + 1 Then 'if there are more files

    Dim bExist As Boolean
    bExist = False
    For i = 0 To lstFile.ListCount
        If FileNameList(intCurFile) = lstFile.List(i) Then
            bExist = True
        End If
    Next 'i
    If bExist = True Then
        Call FileExists(Index) ' if the file exists goto the file exists handler
        Exit Sub
    Else
        FileTransfer(Index).SendData "RQSTFL" & FileNameList(intCurFile) & vbCrLf
    End If
    
    Exit Sub 'exit sub until File has been recived

Else 'Maximum number of files downloaded

    'User.SendData "RQSTFL" & FileNameList(intCurFile) & vbCrLf
    cmdLogin.Enabled = False
    cmdLogin.Caption = "Logged In"
    txtUserName.Enabled = False
    txtPassword.Enabled = False
    cmdProceed.Enabled = True
    'lblstatus(1).Caption = "Logged In To The Server"
    FileTransfer(Index).SendData "DONE" & vbCrLf
    DoEvents
    FileTransfer(Index).Close
    Exit Sub

End If



End Sub
Sub FileExists(Index As Long)
On Error Resume Next
intCurFile = intCurFile + 1 'increase the counter
Call DownloadNextFile(Index)
Exit Sub

End Sub
Public Sub MakeFile(ByVal File As String)
Dim filename As String, filedata As String
On Error Resume Next
'If Left$(File, 4) = "FILE" Then 'check for propper header
    File = Right$(File, Len(File) - 4)  'remove header
    filename = Left$(File, InStr(File, "@") - 1) 'get the filename
    If Left$(filename, 4) = "FILE" Then filename = Mid$(filename, 5, Len(filename))
    If Right$(filename, 3) <> "ini" Then
        filename = App.Path & "\inbox\" & filename
    Else
        filename = App.Path & "\" & filename
    End If
    
    filedata = Mid$(File, InStr(File, "@") + 1, InStr(File, "!") - 1 - InStr(File, "@")) 'get the file data
'    Debug.Assert Len(filename) + Len(filedata) + 2 = Len(File)
    
    '=================================
'    With CommonDialog1
'        .filename = filename
'        .DialogTitle = "Save"
'        .ShowSave
'        filename = .filename
'    End With
    Open filename For Binary Access Write As #1
    Put #1, , filedata
    Close #1
    '=================================
    b64.DecodeFile filename, filename
Close #1
'End If
End Sub
Sub SendDataToAll(strData As String)
On Error Resume Next
For i = 1 To 20
    If Users(i).Enabled = True Then
        Chat(i).SendData strData & vbCrLf
    End If
Next 'i
End Sub

Public Function ConvertNum(Something As String) As String

End Function

Public Function ConvertAlpha(Something As String) As String

End Function

Sub UpdateLadder()
On Error Resume Next
Exit Sub

ladderRefresh = 0
lstRank.Clear
txtHTML.Text = ""
Dim intRanking As Long
intRanking = 0
    Dim userMax As Integer
    Dim nSave As String
    nSave = App.Path & "\user.ini"
    userMax = CInt(GetFromIni("GEN", "TOTAL", nSave))
    For i = 1 To userMax
        Dim curRating As String
        Dim curGetName As String
        curGetName = GetFromIni("USERNUM", CStr(i), nSave)
        curRating = GetFromIni(curGetName, "RATING", nSave)
        
        If curRating = "" Then curRating = "1000"
        
        If curRating < 1000 Then
            badd0 = True
        Else
            badd0 = False
        End If
        Dim curName As String
        curName = curGetName
        
        Dim alphanum(1 To 4) As String
        
        If curRating <> "1000" Then
            If badd0 = True Then 'Less than 1000
                For q = 2 To 4
                    alphanum(q) = Mid(curRating, (q - 1), 1)
                    alphanum(q) = ConvertNum(alphanum(q))
                    DoEvents
                Next ' q
                alphanum(1) = ConvertNum("0")
            End If
            If badd0 = False Then 'More than 1000
                For q = 1 To 4
                alphanum(q) = Mid(curRating, q, 1)
                alphanum(q) = ConvertNum(alphanum(q))
                Next 'q
                DoEvents
            End If
            
            curRating = alphanum(1) & alphanum(2) & alphanum(3) & alphanum(4)
            
            If curRating <> "1000" Then
                lstRank.AddItem curRating & " - " & curName
                DoEvents
            End If
        End If
    Next 'i
    
    Dim tempWins As String
    Dim tempLoss As String
    Dim tempUser As String

    Dim intMaxList As Long
    If lstRank.ListCount > 99 Then
        intMaxList = 99
    Else
        intMaxList = lstRank.ListCount
    End If
    For i = 0 To intMaxList - 1
        'If i <= 15 Then
            tempUser = Mid$(lstRank.List(i), 8, Len(lstRank.List(i)))
            tempWins = GetFromIni(tempUser, "WINS", nSave)
            tempLoss = GetFromIni(tempUser, "LOSS", nSave)
            Call WriteIni(tempUser, "RANK", CStr(i) + 1, nSave)
            strRate = ConvertAlpha(lstRank.List(i))
            lstRank.List(i) = lstRank.List(i) & " - " & tempWins & " - " & tempLoss
            Dim HTMLUser
            HTMLUser = Split(tempUser, " ", Len(tempUser), vbTextCompare)
            Dim TempUser2 As String
            For q = 0 To UBound(HTMLUser)
                If q > 0 Then
                    TempUser2 = "%20" & HTMLUser(q)
                Else
                    TempUser2 = HTMLUser(q)
                End If
            Next 'i
            txtHTML.Text = txtHTML.Text & "<br>" & vbNewLine & "<a href=" & tempUser & ".ini>" & i + 1 & " - " & strRate & " - " & Mid(lstRank.List(i), 8, Len(lstRank.List(i))) & "</a>"
        'ElseIf lstRank.List(i) <> "" Then
        '    strRate = ConvertAlpha(lstRank.List(i))
        '    txtHTML.Text = txtHTML.Text & "<br>" & vbNewLine & "<b>" & i + 1 & "</b> - " & strRate & " - " & Mid(lstRank.List(i), 8, Len(lstRank.List(i)))
        'End If
        DoEvents
    Next 'i
    Dim strIntroMsg As String
    Dim strTime As String
    strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
    
    strIntroMsg = "Welcome to the Ladder Section of Golden Sun: The War of the Adepts.  Here you can find the ranking of the Top 100 Players.  Clicking on a player will display the Win/Loss history of that player.  If the file does not open automatically, open it in a text editor.  Report any absusers on the forums.<br><b>Format: Rank - Rating - Name - Wins - Losses</b>"
    txtHTML.Text = strIntroMsg & vbNewLine & txtHTML.Text & "<br><b>This page is updated approximately every 5 hours.  Last Updated: " & strTime & "</b>"
    Dim FNO As Long
 FNO = FreeFile
 On Error Resume Next
 err.Clear
 Open (strSambarPath & "\ladder.html") For Output As #FNO
  If err.Number <> 0 Then
    'MsgBox "an error has occured"
   Else
    Print #FNO, (txtHTML.Text)
  End If
 Close #FNO
 'On Error GoTo 0
 lstRank.Clear
 txtHTML.Text = ""

End Sub
