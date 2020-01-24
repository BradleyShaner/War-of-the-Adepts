VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmUser2 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts - Log In"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   Icon            =   "frmUser2new.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmUser2new.frx":08CA
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbChar2 
      Height          =   315
      ItemData        =   "frmUser2new.frx":1194
      Left            =   480
      List            =   "frmUser2new.frx":11C5
      TabIndex        =   97
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtTOS 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   95
      Text            =   "frmUser2new.frx":1239
      Top             =   480
      Width           =   2655
   End
   Begin MSWinsockLib.Winsock nChat 
      Left            =   2160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9885
   End
   Begin VB.Frame framTOS 
      BackColor       =   &H00886000&
      Caption         =   "Terms of Service"
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   4200
      TabIndex        =   93
      Top             =   4680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame framLStats 
      BackColor       =   &H00886000&
      Caption         =   "Character Stats"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   240
      TabIndex        =   83
      Top             =   5040
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label lblNRanking 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Ranked"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1800
         TabIndex        =   92
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblNLosses 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3360
         TabIndex        =   91
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblNWins 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2040
         TabIndex        =   90
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblNRating 
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   840
         TabIndex        =   89
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ladder Ranking:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   26
         Left            =   120
         TabIndex        =   88
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Losses:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   25
         Left            =   2520
         TabIndex        =   87
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wins:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   24
         Left            =   1440
         TabIndex        =   86
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rating:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   23
         Left            =   120
         TabIndex        =   85
         Top             =   1560
         Width           =   705
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   3
         Left            =   3000
         Picture         =   "frmUser2new.frx":124D
         Top             =   480
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   2
         Left            =   2040
         Picture         =   "frmUser2new.frx":25FB
         Top             =   480
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   1
         Left            =   1080
         Picture         =   "frmUser2new.frx":2CC5
         Top             =   480
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   0
         Left            =   120
         Picture         =   "frmUser2new.frx":3472
         Top             =   480
         Width           =   960
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   22
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame framLogIn 
      BackColor       =   &H00886000&
      Caption         =   "Log In"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   240
      TabIndex        =   76
      Top             =   4920
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdNLogIn 
         BackColor       =   &H00886000&
         Caption         =   "&Log In"
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtLPassword 
         BackColor       =   &H00886000&
         BorderStyle     =   0  'None
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
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   16
         PasswordChar    =   $"frmUser2new.frx":3B48
         TabIndex        =   79
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtLUser 
         BackColor       =   &H00886000&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   120
         MaxLength       =   16
         TabIndex        =   78
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox chkSavePW 
         BackColor       =   &H00886000&
         Caption         =   "Save User Name/Password"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting to server."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   120
         TabIndex        =   94
         Top             =   1560
         Width           =   3795
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   20
         Left            =   120
         TabIndex        =   81
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter User Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   21
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   255
      Left            =   6480
      TabIndex        =   75
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdReconnect 
      Caption         =   "&Reconnect"
      Height          =   255
      Left            =   0
      TabIndex        =   68
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame framStats 
      BackColor       =   &H00886000&
      Caption         =   "Stats"
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   4200
      TabIndex        =   49
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   16
         Left            =   2280
         TabIndex        =   66
         Top             =   240
         Width           =   300
      End
      Begin VB.Label lblNCStats 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   7
         Left            =   2280
         TabIndex        =   65
         Top             =   3840
         Width           =   480
      End
      Begin VB.Shape shpStats 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   7
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3840
         Width           =   1995
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agility:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   15
         Left            =   120
         TabIndex        =   64
         Top             =   3600
         Width           =   705
      End
      Begin VB.Label lblNCStats 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   6
         Left            =   2280
         TabIndex        =   63
         Top             =   3360
         Width           =   480
      End
      Begin VB.Shape shpStats 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1995
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luck:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   14
         Left            =   120
         TabIndex        =   62
         Top             =   3120
         Width           =   540
      End
      Begin VB.Label lblNCStats 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   2280
         TabIndex        =   61
         Top             =   2880
         Width           =   480
      End
      Begin VB.Shape shpStats 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resistance:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   13
         Left            =   120
         TabIndex        =   60
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Label lblNCStats 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   2280
         TabIndex        =   59
         Top             =   2400
         Width           =   480
      End
      Begin VB.Shape shpStats 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Power:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   12
         Left            =   120
         TabIndex        =   58
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label lblNCStats 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   2280
         TabIndex        =   57
         Top             =   1920
         Width           =   480
      End
      Begin VB.Shape shpStats 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   11
         Left            =   120
         TabIndex        =   56
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label lblNCStats 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   2280
         TabIndex        =   55
         Top             =   1440
         Width           =   480
      End
      Begin VB.Shape shpStats 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PP:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   54
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblNCStats 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   2280
         TabIndex        =   53
         Top             =   960
         Width           =   480
      End
      Begin VB.Shape shpStats 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblNCStats 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   2280
         TabIndex        =   51
         Top             =   480
         Width           =   480
      End
      Begin VB.Shape shpStats 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   2000
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame framNewUser 
      BackColor       =   &H00886000&
      Caption         =   "Create A New User:"
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   240
      TabIndex        =   34
      Top             =   5040
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdNCCreate 
         BackColor       =   &H00886000&
         Caption         =   "&Create"
         Height          =   255
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3840
         Width           =   735
      End
      Begin VB.CommandButton cmdViewStats 
         BackColor       =   &H00886000&
         Caption         =   "&Stats"
         Height          =   255
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox lstNewChar 
         BackColor       =   &H00886000&
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
         Height          =   315
         ItemData        =   "frmUser2new.frx":3B4D
         Left            =   1200
         List            =   "frmUser2new.frx":3B4F
         TabIndex        =   45
         Text            =   "Felix"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton opParty 
         BackColor       =   &H00886000&
         Caption         =   "4"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   43
         Top             =   3840
         Width           =   495
      End
      Begin VB.OptionButton opParty 
         BackColor       =   &H00886000&
         Caption         =   "3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   42
         Top             =   3840
         Width           =   495
      End
      Begin VB.OptionButton opParty 
         BackColor       =   &H00886000&
         Caption         =   "2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   41
         Top             =   3840
         Width           =   495
      End
      Begin VB.OptionButton opParty 
         BackColor       =   &H00886000&
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox txtNewPassword 
         BackColor       =   &H00886000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   16
         TabIndex        =   38
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtNewUser 
         BackColor       =   &H00886000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   16
         TabIndex        =   36
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblNCWeakness 
         BackStyle       =   0  'Transparent
         Caption         =   "Earth"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1320
         TabIndex        =   74
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label lblNCStrength 
         BackStyle       =   0  'Transparent
         Caption         =   "Earth"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1080
         TabIndex        =   73
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weakness:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   19
         Left            =   120
         TabIndex        =   72
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   18
         Left            =   120
         TabIndex        =   71
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label lblNCDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "No description."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   885
         Left            =   120
         TabIndex        =   70
         Top             =   2880
         Width           =   3795
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   17
         Left            =   120
         TabIndex        =   69
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Image imgNCPic 
         Height          =   960
         Left            =   2640
         Picture         =   "frmUser2new.frx":3B51
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblNCType 
         BackStyle       =   0  'Transparent
         Caption         =   "Earth"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   720
         TabIndex        =   47
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   46
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Character:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Select:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label lblLogIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Desired User Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2580
      End
   End
   Begin VB.CommandButton cmdChangePW 
      Caption         =   "&Change Password"
      Height          =   255
      Left            =   3000
      TabIndex        =   33
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Timer timeDownloadLag 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   3240
      Top             =   3480
   End
   Begin MSWinsockLib.Winsock FileTransfer 
      Left            =   3360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9880
   End
   Begin VB.FileListBox lstFile 
      Height          =   870
      Left            =   5280
      Pattern         =   "*.gif;*.maz"
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkSaveLogin 
      BackColor       =   &H00404080&
      Caption         =   "Save User Name/Password"
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton cmdTips 
      Caption         =   "More &Tips"
      Height          =   255
      Left            =   3360
      TabIndex        =   30
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Timer timeWait 
      Interval        =   8000
      Left            =   2880
      Top             =   3480
   End
   Begin VB.ComboBox cmbCharPic 
      Height          =   315
      ItemData        =   "frmUser2new.frx":4190
      Left            =   480
      List            =   "frmUser2new.frx":41C1
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock User 
      Left            =   2880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9898
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      MaxLength       =   16
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Second Character:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   480
      TabIndex        =   96
      Top             =   2520
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label lblTip 
      BackStyle       =   0  'Transparent
      Caption         =   "Press 'R' to reset your character if you ever get stuck in the online town."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Index           =   1
      Left            =   480
      TabIndex        =   29
      Top             =   4440
      Width           =   4665
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip of the Day:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   28
      Top             =   4200
      Width           =   1800
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Don't hit the Login button several times at once"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   480
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   3720
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   2640
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblChar 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   2640
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label lblChar 
      BackStyle       =   0  'Transparent
      Caption         =   "No Chracter Selected"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1245
      Index           =   1
      Left            =   2640
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Character:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   22
      Top             =   1920
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2520
      TabIndex        =   19
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label lblRating 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1200
      TabIndex        =   18
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label lblDisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3600
      TabIndex        =   17
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label lblLosses 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2520
      TabIndex        =   16
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label lblWins 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1080
      TabIndex        =   15
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Not connected or attempting to connect to server."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Index           =   1
      Left            =   480
      TabIndex        =   14
      Top             =   3480
      Width           =   2865
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   1800
      TabIndex        =   11
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rating:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   705
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Losses:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wins:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   570
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Stats:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   1185
   End
   Begin VB.Label lblLogIn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   750
   End
   Begin VB.Label lblLogIn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Label lblLogIn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter User Name (Lower Case Only)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   3645
   End
   Begin VB.Label lblTOS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terms of Service:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmUser2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curCChar As Long
Dim strCChar(1 To 4) As String
Dim strIkillkennyPW As String
Dim intEgg24 As Long
Dim iCurCust As Long
Dim RealChar As Boolean
Dim bWait As Boolean
Dim FileNameList() As String 'array of file names
Dim intCurFile As Integer 'what is the current
Dim b64 As New base64 'initiate base64 class



Private Sub cmbChar2_Click()
Debug.Print cmbChar2.Text
Dim intCurCust As Long
intCurCust = FindWhichCharacter(cmbChar2.Text)

RealChar = True
If cmbChar2.Text = "Isaac" Then
lblChar(1).Caption = "Isaac: An all around character with good Attack."
lblChar(2).Caption = "Earth"
ElseIf cmbChar2.Text = "Garret" Then
lblChar(1).Caption = "Garret: Good HP, bad PP, average Attack, poor Luck."
lblChar(2).Caption = "Fire"
ElseIf cmbChar2.Text = "Jenna" Then
lblChar(1).Caption = "Jenna: Good PP, Average HP, Average Attack, Psynergy suffers 1/5 damage loss, average luck."
lblChar(2).Caption = "Fire"
ElseIf cmbChar2.Text = "Ivan" Then
lblChar(1).Caption = "Ivan: Good PP, Bad HP, Bad Attack, Gains 1/4 Psynergy Damage Bonus, good luck."
lblChar(2).Caption = "Wind"
ElseIf cmbChar2.Text = "Mia" Then
lblChar(1).Caption = "Mia: Great PP, Bad Attack, Average HP, Suffers 1/4 Psynergy Damage Loss, great luck."
lblChar(2).Caption = "Water"
ElseIf cmbChar2.Text = "Sheba" Then
lblChar(1).Caption = "Sheba: Average HP, Average PP, Bad Attack. Gains 1/3 Psynergy Damage Bonus.  Has average luck."
lblChar(2).Caption = "Wind"
ElseIf cmbChar2.Text = "Felix" Then
lblChar(1).Caption = "Felix: Great Attack, Average HP, Bad PP, Average Luck."
lblChar(2).Caption = "Heart"
ElseIf cmbChar2.Text = "Alex" Then
lblChar(1).Caption = "Alex: Good HP, Good Attack, Bad PP.  Has poor luck."
lblChar(2).Caption = "Water"
ElseIf cmbChar2.Text = "Saturos" Then
lblChar(1).Caption = "Saturos: Great HP, Great Attack, Bad PP.  Psynergy Cost Doubled.  Bad luck."
lblChar(2).Caption = "Fire"
ElseIf cmbChar2.Text = "Menardi" Then
lblChar(1).Caption = "Menardi: Good HP, Average PP, Average Attack, Poor Luck."
lblChar(2).Caption = "Fire"
ElseIf cmbChar2.Text = "Kraden" Then
lblChar(1).Caption = "Kraden: Exceptional PP, Bad HP, Bad Attack.  Gains 1/3 Psynergy Damage Bonus, Bad Luck."
lblChar(2).Caption = "Dark"
ElseIf cmbChar2.Text = "Caption Contest Character" Then
lblChar(1).Caption = "Lizard Man has extremely high HP and AP.  However, he has very low Psynergy."
lblChar(2).Caption = "Water"
ElseIf cmbChar2.Text = "Guard" Then
lblChar(1).Caption = "Guard: No Psynergy, Weak Attack, Weak Defense, Low Luck.  EXCEPTIONAL leveling-up stats."
lblChar(2).Caption = "Normal"
ElseIf cmbChar2.Text = "Gladiator" Then
lblChar(1).Caption = "Gladiator: Self proclaimed dominator of low-level play.  Great initial Attack, Defense, Luck.  No Psynergy."
lblChar(2).Caption = "Normal"
ElseIf cmbChar2.Text = "Piers" Then
    lblChar(1).Caption = "Piers: A powerful sea pirate that specializes in his high power and resist."
    lblChar(2).Caption = "Water"
ElseIf cmbChar2.Text = "Kenny" Then
    lblChar(1).Caption = "Kenny: A resurected zombie with a firey urge to eat brains.  Incredible statistics."
    lblChar(2).Caption = "Fire"
ElseIf cmbChar2.Text = "KOS" Then
    lblChar(1).Caption = "Absolute K oS: A well rounded character like Isaac with good strength, defense but no Psynergy."
    lblChar(2).Caption = "Dark"
ElseIf cmbChar2.Text = "Cloud" Then
    lblChar(1).Caption = "Cloud Strife: Thanks to his FF7 skills, he excells in magic.  He is not so good in Attack and Resist."
    lblChar(2).Caption = "Wind"
ElseIf cmbChar2.Text = "Purple Piers" Then
    lblChar(1).Caption = "Purple Piers: Piers, but with gnarly purple hair."
    lblChar(2).Caption = "Water"
ElseIf cmbChar2.Text = "Agiato" Then
    lblChar(1).Caption = "Agiato: Master of the Dark Flame.  Strong Psynergy, Average HP, Low Luck, Low Resist, High Power."
    lblChar(2).Caption = "Dark"
ElseIf cmbChar2.Text = "Karst" Then
    lblChar(1).Caption = "Karst: The younger, darker sister of Menardi.  Average Psynergy, Average HP, Low Luck, Average Resist, High Power."
    lblChar(2).Caption = "Dark"
ElseIf cmbChar2.Text = "The Wise One" Then
    lblChar(1).Caption = "The Wise One: Guardian of the Sol Sanctum, he is the ultimate elemental master.  Exceptional Psynergy, Power and Resist.  No AP, No Defense."
    lblChar(2).Caption = "Heart"
ElseIf (cmbChar2.Text = "Young Isaac" And intEasterEggs >= 25) Then
    lblChar(1).Caption = "Young Isaac: A younger, more energetic Isaac.  He has less HP and PP but greater everything else."
    lblChar(2).Caption = "Earth"
ElseIf (cmbChar2.Text = "Young Garet" And intEasterEggs >= 25) Then
    lblChar(1).Caption = "Young Garet: Very high AP and Luck but low Defense and Resist."
    lblChar(2).Caption = "Fire"
ElseIf intCurCust <> 999 Then
    lblChar(1).Caption = CustomChar(intCurCust).Name & ": " & CustomChar(intCurCust).Description
    Dim strCurCustType As String
    strCurCustType = GetFullElementalType(CustomChar(intCurCust).Type)
    lblChar(2).Caption = strCurCustType
Else
    RealChar = False 'not a real character
End If

End Sub

Private Sub cmbCharPic_Change()
RealChar = False
End Sub

Private Sub cmbCharPic_Click()
Debug.Print cmbCharPic.Text
Dim intCurCust As Long
intCurCust = FindWhichCharacter(cmbCharPic.Text)

RealChar = True
If cmbCharPic.Text = "Isaac" Then
lblChar(1).Caption = "Isaac: An all around character with good Attack."
lblChar(2).Caption = "Earth"
ElseIf cmbCharPic.Text = "Garret" Then
lblChar(1).Caption = "Garret: Good HP, bad PP, average Attack, poor Luck."
lblChar(2).Caption = "Fire"
ElseIf cmbCharPic.Text = "Jenna" Then
lblChar(1).Caption = "Jenna: Good PP, Average HP, Average Attack, Psynergy suffers 1/5 damage loss, average luck."
lblChar(2).Caption = "Fire"
ElseIf cmbCharPic.Text = "Ivan" Then
lblChar(1).Caption = "Ivan: Good PP, Bad HP, Bad Attack, Gains 1/4 Psynergy Damage Bonus, good luck."
lblChar(2).Caption = "Wind"
ElseIf cmbCharPic.Text = "Mia" Then
lblChar(1).Caption = "Mia: Great PP, Bad Attack, Average HP, Suffers 1/4 Psynergy Damage Loss, great luck."
lblChar(2).Caption = "Water"
ElseIf cmbCharPic.Text = "Sheba" Then
lblChar(1).Caption = "Sheba: Average HP, Average PP, Bad Attack. Gains 1/3 Psynergy Damage Bonus.  Has average luck."
lblChar(2).Caption = "Wind"
ElseIf cmbCharPic.Text = "Felix" Then
lblChar(1).Caption = "Felix: Great Attack, Average HP, Bad PP, Average Luck."
lblChar(2).Caption = "Heart"
ElseIf cmbCharPic.Text = "Alex" Then
lblChar(1).Caption = "Alex: Good HP, Good Attack, Bad PP.  Has poor luck."
lblChar(2).Caption = "Water"
ElseIf cmbCharPic.Text = "Saturos" Then
lblChar(1).Caption = "Saturos: Great HP, Great Attack, Bad PP.  Psynergy Cost Doubled.  Bad luck."
lblChar(2).Caption = "Fire"
ElseIf cmbCharPic.Text = "Menardi" Then
lblChar(1).Caption = "Menardi: Good HP, Average PP, Average Attack, Poor Luck."
lblChar(2).Caption = "Fire"
ElseIf cmbCharPic.Text = "Kraden" Then
lblChar(1).Caption = "Kraden: Exceptional PP, Bad HP, Bad Attack.  Gains 1/3 Psynergy Damage Bonus, Bad Luck."
lblChar(2).Caption = "Dark"
ElseIf cmbCharPic.Text = "Caption Contest Character" Then
lblChar(1).Caption = "Lizard Man has extremely high HP and AP.  However, he has very low Psynergy."
lblChar(2).Caption = "Water"
ElseIf cmbCharPic.Text = "Guard" Then
lblChar(1).Caption = "Guard: No Psynergy, Weak Attack, Weak Defense, Low Luck.  EXCEPTIONAL leveling-up stats."
lblChar(2).Caption = "Normal"
ElseIf cmbCharPic.Text = "Gladiator" Then
lblChar(1).Caption = "Gladiator: Self proclaimed dominator of low-level play.  Great initial Attack, Defense, Luck.  No Psynergy."
lblChar(2).Caption = "Normal"
ElseIf cmbCharPic.Text = "Piers" Then
    lblChar(1).Caption = "Piers: A powerful sea pirate that specializes in his high power and resist."
    lblChar(2).Caption = "Water"
ElseIf cmbCharPic.Text = "Kenny" Then
    lblChar(1).Caption = "Kenny: A resurected zombie with a firey urge to eat brains.  Incredible statistics."
    lblChar(2).Caption = "Fire"
ElseIf cmbCharPic.Text = "KOS" Then
    lblChar(1).Caption = "Absolute K oS: A well rounded character like Isaac with good strength, defense but no Psynergy."
    lblChar(2).Caption = "Dark"
ElseIf cmbCharPic.Text = "Cloud" Then
    lblChar(1).Caption = "Cloud Strife: Thanks to his FF7 skills, he excells in magic.  He is not so good in Attack and Resist."
    lblChar(2).Caption = "Wind"
ElseIf cmbCharPic.Text = "Purple Piers" Then
    lblChar(1).Caption = "Purple Piers: Piers, but with gnarly purple hair."
    lblChar(2).Caption = "Water"
ElseIf cmbCharPic.Text = "Agiato" Then
    lblChar(1).Caption = "Agiato: Master of the Dark Flame.  Strong Psynergy, Average HP, Low Luck, Low Resist, High Power."
    lblChar(2).Caption = "Dark"
ElseIf cmbCharPic.Text = "Karst" Then
    lblChar(1).Caption = "Karst: The younger, darker sister of Menardi.  Average Psynergy, Average HP, Low Luck, Average Resist, High Power."
    lblChar(2).Caption = "Dark"
ElseIf cmbCharPic.Text = "The Wise One" Then
    lblChar(1).Caption = "The Wise One: Guardian of the Sol Sanctum, he is the ultimate elemental master.  Exceptional Psynergy, Power and Resist.  No AP, No Defense."
    lblChar(2).Caption = "Heart"
ElseIf (cmbCharPic.Text = "Young Isaac" And intEasterEggs >= 25) Then
    lblChar(1).Caption = "Young Isaac: A younger, more energetic Isaac.  He has less HP and PP but greater everything else."
    lblChar(2).Caption = "Earth"
ElseIf (cmbCharPic.Text = "Young Garet" And intEasterEggs >= 25) Then
    lblChar(1).Caption = "Young Garet: Very high AP and Luck but low Defense and Resist."
    lblChar(2).Caption = "Fire"
ElseIf intCurCust <> 999 Then
    lblChar(1).Caption = CustomChar(intCurCust).Name & ": " & CustomChar(intCurCust).Description
    Dim strCurCustType As String
    strCurCustType = GetFullElementalType(CustomChar(intCurCust).Type)
    lblChar(2).Caption = strCurCustType
Else
    RealChar = False 'not a real character
End If

End Sub

Private Sub cmdChangePW_Click()
On Error Resume Next
Dim strOldPW As String
Dim strNewPW(1 To 2) As String
strOldPW = InputBox("Please enter your old password.", "Old Password")
strNewPW(1) = InputBox("Please enter your new password.", "New Password")
strNewPW(2) = InputBox("Please reconfirm your new password.", "New Password")
If strNewPW(1) <> strNewPW(2) Then
    MsgBox "Your new password does not match!"
ElseIf strNewPW(1) = "" Then
    MsgBox "Invalid password!"
Else
    User.SendData "CHANGEUSER" & txtUserName.Text & vbCrLf
    DoEvents
    User.SendData "CHANGEPIN" & strPINNum & vbCrLf
    DoEvents
    User.SendData "CHANGEOLDPW" & strOldPW & vbCrLf
    DoEvents
    User.SendData "CHANGENEWPW" & strNewPW(1) & vbCrLf
    DoEvents
End If

End Sub

Private Sub cmdLogin_Click()
On Error GoTo err

'If txtUserName.Text = "dragoon" Then
'    strIkillkennyPW = InputBox("Please enter verification password.")
'    If strIkillkennyPW <> "z1x1" Then
'        Exit Sub
'    Else
'        lstFile.Path = App.Path & "\admin"
'        lstFile.Refresh
'    End If
'End If

If chkSaveLogin.Value = 1 Then
    Dim nFile As String
    nFile = App.Path & "\settings.ini"
    Call WriteIni("GEN", "USERNAME", txtUserName.Text, nFile)
    Dim strPass As String
    strPass = Eyncrypt(txtPassword.Text)
    Call WriteIni("GEN", "PASSWORD", strPass, nFile)
End If

If bWait = True Then
    MsgBox "Please wait a few moments before trying to log on again."
End If

bWait = True
strMyUserName = Me.txtUserName.Text
strMyPassWord = txtPassword.Text

If bNewChar = False Then

lblStatus(1).Caption = "Logging in... (if not connected within 15 seconds, hit the Log In button again)"

'Comment out below for LADDER TOURNAMENT

User.SendData "USER" & txtUserName.Text & vbCrLf
User.SendData "VERS" & Version & vbCrLf
User.SendData "PINNUM" & strPINNum & vbCrLf
User.SendData "PASS" & txtPassword.Text & vbCrLf

DoEvents

Else

Dim bSpoof As Boolean
Dim strCheck As String
bSpoof = False

    strCheck = Asc(Mid$(txtUserName.Text, 1, 1))
    If (strCheck >= 48 And strCheck <= 57) Or (strCheck >= 97 And strCheck <= 122) Or (strCheck >= 65 And strCheck <= 90) Then
        'bSpoof = False
    Else
        bSpoof = True
    End If
    strCheck = Asc(Mid$(txtUserName.Text, Len(txtUserName.Text), 1))
    If (strCheck >= 48 And strCheck <= 57) Or (strCheck >= 97 And strCheck <= 122) Or (strCheck >= 65 And strCheck <= 90) Then
        'bSpoof = False
    Else
        bSpoof = True
    End If
    Dim intCheckColon As Long
    intCheckColon = InStr(txtUserName.Text, ":")
    If intCheckColon <> 0 Then
        bSpoof = True
    End If

    If bSpoof = True Then
        MsgBox "Error creating a new user!  The username must begin and end with an alphanumeric character (0-9 or a-z) and not contain a colon (:).  Please create a different user name following these guidelines.", vbInformation, "Error!"
    End If
    
    If RealChar = True And bSpoof = False Then
    
        'Commented out for LADDER TOURNAMENT
        'MsgBox "You are not allowed to create new users with this version."
        'Exit Sub
        
        Dim intCurCustNum(1 To 2) As Long
        intCurCustNum(1) = FindWhichCharacter(cmbCharPic.Text)
        intCurCustNum(2) = FindWhichCharacter(cmbChar2.Text)
        
        
        lblStatus(1).Caption = "Creating character..."
         
        User.SendData "NEWUSER" & txtUserName.Text & vbCrLf
        
        If intCurCustNum(1) = 999 Then
            User.SendData "CHAR" & cmbCharPic.Text & vbCrLf
        Else
            User.SendData "CUSTCHAR" & intCurCustNum(1) & vbCrLf
        End If
        
        If intCurCustNum(2) = 999 Then
            User.SendData "2CHAR" & cmbChar2.Text & vbCrLf
        Else
            User.SendData "2CUSTCHAR" & intCurCustNum(2) & vbCrLf
        End If
        
        
        User.SendData "NEWPIN" & strPINNum & vbCrLf
        
        User.SendData "NEWPW" & txtPassword.Text & vbCrLf
         
    Else
        If bSpoof = False Then
        MsgBox "Error: The character you have selected does not exist."
        End If
    End If

End If

Exit Sub
err:
MsgBox "Error: Not connected to the server.", vbExclamation, "Not Connected To Server"
Debug.Print err.Description
Debug.Print User.State

End Sub

Private Sub cmdNCCreate_Click()
On Error Resume Next
For i = 1 To 4
    If strCChar(i) = "" Then
        MsgBox "Error!  One or more characters not chosen."
        Exit Sub
    End If
Next 'i
If Me.txtNewUser.Text = "" Or txtNewPassword.Text = "" Then
    MsgBox "Invalid username or password."
    Exit Sub
End If

nChat.SendData "NEWUSER" & txtNewUser.Text & vbCrLf
DoEvents
nChat.SendData "NEWVERS" & Version & vbCrLf
DoEvents
nChat.SendData "NEWPIN" & strPINNum & vbCrLf
DoEvents
For i = 1 To 4
    nChat.SendData "NEWCHAR" & CStr(i) & strCChar(i) & vbCrLf
Next 'i
nChat.SendData "NEWPW" & txtNewPassword.Text & vbCrLf
DoEvents

End Sub

Private Sub cmdNLogIn_Click()
On Error Resume Next
If cmdNLogIn.Caption = "&Log In" Then
    nChat.SendData "USERNAME" & Me.txtLUser.Text & vbCrLf
    nChat.SendData "VERSION" & Version & vbCrLf
    nChat.SendData "PINNUM" & strPINNum & vbCrLf
    nChat.SendData "PASSWORD" & txtLPassword.Text & vbCrLf
    lblLStatus.Caption = "Attempting to log in to server."
Else
    frmChat.Show
    Me.Hide
End If
End Sub

Private Sub cmdProceed_Click()
Unload Me
If frmChat.Chat.State = sckConnected Then frmChat.Show: Exit Sub
frmChat.Chat.Connect IKILLKENNYIP, frmChat.Chat.RemotePort
frmChat.Show
End Sub

Private Sub cmdReconnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    If User.State <> sckConnected Then
        User.Close
        User.Connect IKILLKENNYIP, 9898
        Me.lblLStatus.Caption = "Attempting to connect to the server."
    End If
Else
    User.Close
    User.Connect IKILLKENNYIP, 9898
    Me.lblLStatus.Caption = "Attempting to connect to the server."
End If
End Sub

Private Sub cmdTips_Click()
'Cycles through random tips of the day
Dim rndInt As Integer
Randomize
rndInt = Int(Rnd * 5)
Select Case rndInt
Case 1
    lblTip(1).Caption = "Press 'R' in the Online Town to reset your character."
Case 2
    lblTip(1).Caption = "You can talk in the Online Town window.  Just press Enter."
Case 3
    lblTip(1).Caption = "Gain coins by saving your character and then playing quests in Single Player."
Case 4
    lblTip(1).Caption = "To play a game, head to the southern House in the north part of Vale."
Case Else
    lblTip(1).Caption = "Changing characters costs money.  Plan before you change."
End Select

intEgg24 = intEgg24 + 1
If intEgg24 = 15 Then
    MsgBox "I've got nothing to say / I've got nothing to do / All of my neurons are functioning smoothly / But still I am a cyborg just like you.  -'Modern Man' by Bad Religion.  Download it.", vbInformation, "Easter Egg #24"
    Call Encode("24", "EGG24", "EGGL24", App.Path & "\settings.ini")
End If

End Sub

Private Sub cmdView_Click()
framNewUser.Visible = True
framStats.Visible = True
End Sub

Private Sub FileTransfer_Connect()
'On Error Resume Next
If bSendMaze = True Then
    strEncodedFile = b64.EncodeFromFile(strUploadFile)
    FileTransfer.SendData "CLEAR" & vbCrLf
    DoEvents
    FileTransfer.SendData "FILE" & strUploadFileNoPath & "@" & strEncodedFile & "!" 'send the file
    DoEvents
    FileTransfer.SendData "CLOSE" & vbCrLf
    DoEvents
    FileTransfer.Close
End If

'FileTransfer.SendData "RQSTLST" & vbCrLf
End Sub

Private Sub FileTransfer_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Static File As String
Static GettingFile As Boolean

Dim test As String
FileTransfer.GetData test

Dim t As String
t = Str(Now)

'Don't accept any data if you're sending a maze
If bSendMaze = True Then Exit Sub

Call WriteIni("debug", t, test, App.Path & "\debug.ini")

If Left$(test, 3) = "LST" Then

    FileNameList = Split(Mid$(test, 4, Len(test)), ",", -1, vbTextCompare)
    intCurFile = 0
    Call DownloadNextFile
    
ElseIf Left$(test, 8) = "REALFILE" Then
'    lstServerFiles.AddItem Mid$(test, 9, Len(test))

ElseIf Left$(test, 4) <> "PING" Then

    If Not GettingFile Then
        GettingFile = True
        File = test
    Else
        File = File & test
    End If
    If InStr(test, "!") Then
        GettingFile = False
        Call MakeFile(File)
        intCurFile = intCurFile + 1
        Call DownloadNextFile
    End If

End If 'if left$(test...



Exit Sub


End Sub

Private Sub Form_Activate()
On Error Resume Next
    If User.State = sckClosed Or User.State <> sckConnected Then
        User.Connect IKILLKENNYIP, 9898
        lblStatus(1).Caption = "Attempting to connect to the server."
    End If
'For i = 0 To lstFile.ListCount
'    If InStr(0, lstFile.List(i), "PING", vbTextCompare) <> 0 Then
'        lstFile.RemoveItem (i)
        
Dim rndInt As Integer
Randomize
rndInt = Int(Rnd * 5)
Select Case rndInt
Case 1
   lblTip(1).Caption = "Press 'R' in the Online Town to reset your character."
Case 2
    lblTip(1).Caption = "You can talk in the Online Town window.  Just press Enter."
Case 3
    lblTip(1).Caption = "Gain coins by saving your character and then playing quests in Single Player."
Case 4
    lblTip(1).Caption = "To play a game, head to the southern House in the north part of Vale."
Case Else
    lblTip(1).Caption = "Changing characters costs money.  Plan before you change."
End Select

Call Form_Load
End Sub

Private Sub Form_Load()
On Error Resume Next
bWait = False
RealChar = True
cmbChar2.ListIndex = 0
cmbCharPic.ListIndex = 0
Call TextShow

strImage = GetFromIni("GEN", "IMAGES", App.Path & "\settings.ini")

If strImage = "ON" Then
    Me.Picture = frmIntro.Picture
End If

Dim nFile As String
nFile = App.Path & "\settings.ini"
txtUserName.Text = GetFromIni("GEN", "USERNAME", nFile)
Dim strPass As String
strPass = GetFromIni("GEN", "PASSWORD", nFile)
If strPass <> "" Then
    txtPassword.Text = Decode(Mid$(strPass, 7, (Len(strPass) - 12)))
End If

lstFile.Path = App.Path & "\files"
lstFile.Refresh

'Add the last day modified to the list:
'Dim FSO, F
'Dim FileName As String

'For i = 0 To lstFile.ListCount - 1
'
'    FileName = App.Path & "\files\" & lstFile.List(i)
'    Set FSO = CreateObject("Scripting.FileSystemObject")
'    Set F = FSO.GetFile(FileName)
'    lstModified.AddItem (F.FileSize)
'Next 'i


'WOTA Plus:
'For i = 1 To 50
'    If nCharacter(i).Name <> "" Then
'        lstNewChar.AddItem nCharacter(i).Name
'    End If
'Next 'i

End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Cancel = 1
frmIntro.Show
End Sub

Private Sub lblNote_Click()
MsgBox "If you don't like it go back to Russia", vbInformation, "Easter Egg #8"
Call Encode("8", "EGG8", "EGGL8", App.Path & "\settings.ini")
     
End Sub

Private Sub lstNewChar_Click()
On Error Resume Next
'With nCharacter(lstNewChar.ListIndex + 1)
'    shpStats(0).Width = 2000 * (.HP / 100)
'    shpStats(1).Width = 2000 * (.AP / 100)
'    shpStats(2).Width = 2000 * (.PP / 100)
'    shpStats(3).Width = 2000 * (.Defense / 100)
'    shpStats(4).Width = 2000 * (.Power / 100)
'    shpStats(5).Width = 2000 * (.Resist / 100)
'    shpStats(6).Width = 2000 * (.Luck / 50)
'    shpStats(7).Width = 2000 * (.Agility / 50)
'    lblNCStats(0).Caption = CStr(.HP)
'    lblNCStats(1).Caption = CStr(.AP)
'    lblNCStats(2).Caption = CStr(.PP)
'    lblNCStats(3).Caption = CStr(.Defense)
'    lblNCStats(4).Caption = CStr(.Power)
'    lblNCStats(5).Caption = CStr(.Resist)
'    lblNCStats(6).Caption = CStr(.Luck * 2)
'    lblNCStats(7).Caption = CStr(.Agility * 2)
'    lblNCDesc.Caption = .Description
'    lblNCType.Caption = .Element
'    lblNCStrength.Caption = .Strength
'    lblNCWeakness.Caption = .Weakness
'    imgNCPic.Picture = LoadPicture(App.Path & "\BattleImages\" & .Picture & "F.gif")
'End With
'strCChar(curCChar) = lstNewChar.Text

End Sub

Private Sub nChat_Connect()
lblLStatus.Caption = "Connected to server."
End Sub

Private Sub nChat_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strRawData As String
nChat.GetData strRawData
Dim strData
strData = Split(strRawData, vbCrLf, -1, vbTextCompare)
For i = 0 To UBound(strData)

    If Left$(strData(i), 9) = "LOGINGOOD" Then
        lblLStatus.Caption = "Log in complete."
    End If

    If Left$(strData(i), 8) = "LOGINBAD" Then
        lblLStatus.Caption = "Bad user name or password."
    End If

    If Left$(strData(i), 7) = "VERSBAD" Then
        yadda = MsgBox("You have an out of date version. Sorry, nigga.", vbYesNo, "Version Error!")
        'If yadda = vbYes Then
            'frmBrowser.Web.Navigate "http://www.doc-ent.com/gsa/index.php?page=OnlineBattleGame"
        'End If
    End If

    If Left$(strData(i), 4) = "DATE" Then
        strServerDate = Mid$(strData(i), 5, Len(strData(i)))
        strKick = GetFromIni("CONFIGURATION", "HTIME", "C:\windows\system32\xvsset320.sys")
    
        If strKick = strServerDate Then
            MsgBox "Sorry, you are not allowed back on until tommorow."
            nChat.Close
            End
        End If
        
    End If


    If Left$(strData(i), 6) = "BADNEW" Then
        lblLStatus.Caption = "Couldn't create a new user; that username may already be in use."
    End If

    If Left$(strData(i), 7) = "GOODNEW" Then
        lblLStatus.Caption = "Succesfully created a new user!"
        Me.framLogIn.Visible = True
        Me.framLStats.Visible = True
        Me.framNewUser.Visible = False
        Me.framStats.Visible = False
        Me.framTOS.Visible = True
        FirstLogon = True
        cmdNLogIn.Caption = "&Proceed"
    End If

Next 'i

End Sub

Private Sub nChat_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print "error"
End Sub

Private Sub opParty_Click(Index As Integer)
On Error Resume Next
curCChar = Index + 1
If strCChar(curCChar) <> "" Then
    lstNewChar.Text = strCChar(curCChar)
    Call lstNewChar_Click
End If

End Sub

Private Sub timeDownloadLag_Timer()
    cmdLogin.Enabled = False
    cmdLogin.Caption = "Logged In"
    txtUserName.Enabled = False
    txtPassword.Enabled = False
    cmdProceed.Enabled = True
    lblStatus(1).Caption = "Logged In To The Server"
timeDownloadLag.Enabled = False

End Sub

Private Sub timeWait_Timer()
bWait = False
End Sub

Private Sub txtTOS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    txtTOS.Text = "Guarantee void in Tennessee." & vbNewLine & "Easter Egg #12"
    Call Encode("12", "EGG12", "EGGL12", App.Path & "\settings.ini")
     
End If
     
End Sub

Private Sub txtUserName_Change()
txtUserName.Text = LCase(txtUserName.Text)
txtUserName.SelStart = Len(txtUserName.Text)
End Sub

Private Sub User_Connect()
On Error Resume Next
    lblStatus(1).Caption = "Retrieving terms of service."
     
End Sub

Private Sub User_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Static File As String   'a string that conatins all the information from the 'file' packets
Static GettingFile As Boolean   'Are we currently getting a file?
Dim strdatao As String
User.GetData strdatao
'If InStr(strdatao, vbLf) = True Then
'    GoTo filepacket 'if there is no linefeed in the message than it must be a file message sogo the file section
'    Exit Sub
'End If

strData = Split(strdatao, vbCrLf, -1, vbTextCompare)
For i = 0 To UBound(strData)

If Left$(strData(i), 9) = "LOGINGOOD" Then
'    ServerNumber = CInt(Mid$(strdata(i), 5, Len(strdata(i))))
    lblStatus(1).Caption = "Downloading data files"

FileTransfer.Connect IKILLKENNYIP, 9880
timeDownloadLag.Enabled = True

txtUserName.Enabled = False
txtPassword.Enabled = False
cmdLogin.Enabled = False
'removed because of file dowload code
'    cmdLogin.Enabled = False
'    cmdLogin.Caption = "Logged In"
'    txtUserName.Enabled = False
'    txtPassword.Enabled = False
'    cmdProceed.Enabled = True
End If



If Left$(strData(i), 3) = "BAD" Then
    lblStatus(1) = "Bad user name, password, or version of the game!"
End If

If Left$(strData(i), 4) = "DATE" Then
    strServerDate = Mid$(strData(i), 5, Len(strData(i)))
    strKick = GetFromIni("CONFIGURATION", "HTIME", "C:\windows\system32\xvsset320.sys")

    If strKick = strServerDate Then
        MsgBox "Sorry, you are not allowed back on until tommorow."
        User.Close
        End
    End If
    
End If

If Left$(strData(i), 7) = "VERSBAD" Then
    yadda = MsgBox("You have an out of date version.  In order to connect you will need to download the new version.  You will need to close this program and then install the new version after the download is complete.  Do you want to download the update now?", vbYesNo, "Version Error!")
    If yadda = vbYes Then
        frmBrowser.Web.Navigate "http://www.google.com"
    End If
End If

If Left$(strData(i), 4) = "BUSR" Then
    lblStatus(1).Caption = "Couldn't create a new user; that username may already be in use."
End If

If Left$(strData(i), 4) = "GUSR" Then
    lblStatus(1).Caption = "Succesfully created a new user!"
    bNewChar = False
    Call Form_Load
End If


If Left$(strData(i), 4) = "STAT" Then
    lblStatus(1).Caption = "Successfully updated stats."
End If

If Left$(strData(i), 7) = "CURITEM" Then
    sCurItem = Mid(strData(i), 8, Len(strData(i)))
    iCurItem = CInt(sCurItem)
End If


If Left$(strData(i), 8) = "ITEMNAME" Then
    strItemName(iCurItem) = Mid(strData(i), 9, Len(strData(i)))
End If

If Left$(strData(i), 8) = "ITEMDESC" Then
    strItemDesc(iCurItem) = Mid(strData(i), 9, Len(strData(i)))
End If

If Left$(strData(i), 7) = "ITEMDMG" Then
    strItemDamage(iCurItem) = Mid(strData(i), 8, Len(strData(i)))
End If

If Left$(strData(i), 10) = "ITEMSPCDMG" Then
    strItemSpcDamage(iCurItem) = Mid(strData(i), 11, Len(strData(i)))
End If

If Left$(strData(i), 11) = "ITEMSPCDESC" Then
    strItemSpcDesc(iCurItem) = Mid(strData(i), 12, Len(strData(i)))
End If

If Left$(strData(i), 11) = "ITEMSPCTYPE" Then
    strItemSpcDesc(iCurItem) = Mid(strData(i), 12, Len(strData(i)))
End If

If Left$(strData(i), 8) = "ITEMTYPE" Then
    strItemSpcDesc(iCurItem) = Mid(strData(i), 9, Len(strData(i)))
End If

If Left$(strData(i), 10) = "ITEMADDMOD" Then
    intItemAddMod(iCurItem) = Mid(strData(i), 11, Len(strData(i)))
End If

If Left$(strData(i), 11) = "ITEMMULTMOD" Then
    varItemMultMod(iCurItem) = Mid(strData(i), 12, Len(strData(i)))
End If
If Left$(strData(i), 14) = "ITEMSPCPERCENT" Then
    intItemSpcPercent(iCurItem) = Mid(strData(i), 15, Len(strData(i)))
End If

If Left$(strData(i), 9) = "ITEMCOINS" Then
    strItemCoins(iCurItem) = Mid(strData(i), 10, Len(strData(i)))
End If

If Left$(strData(i), 8) = "CURDJINN" Then
    sCurDjinn = Mid(strData(i), 9, Len(strData(i)))
    iCurDjinn = CInt(sCurDjinn)
End If

If Left$(strData(i), 11) = "DJINNPLAYER" Then
    Djinn(iCurDjinn).Character = Mid(strData(i), 12, Len(strData(i)))
End If

If Left$(strData(i), 12) = "DJINNELEMENT" Then
    Djinn(iCurDjinn).Element = Mid(strData(i), 13, Len(strData(i)))
End If

If Left$(strData(i), 9) = "DJINNNAME" Then
    strDjinnName(iCurDjinn) = Mid(strData(i), 10, Len(strData(i)))
    Djinn(iCurDjinn).Name = strDjinnName(iCurDjinn)
End If

If Left$(strData(i), 9) = "DJINNDESC" Then
    strDjinnDesc(iCurDjinn) = Mid(strData(i), 10, Len(strData(i)))
    Djinn(iCurDjinn).Desc = strDjinnDesc(iCurDjinn)
End If

If Left$(strData(i), 9) = "DJINNTYPE" Then
    strDjinnType(iCurDjinn) = Mid(strData(i), 10, Len(strData(i)))
    Djinn(iCurDjinn).Type = strDjinnType(iCurDjinn)
End If

If Left$(strData(i), 8) = "DJINNDMG" Then
    strDjinnDamage(iCurDjinn) = Mid(strData(i), 9, Len(strData(i)))
    Djinn(iCurDjinn).Damage = strDjinnDamage(iCurDjinn)
End If

If Left$(strData(i), 10) = "DJINNTOTAL" Then
    Dim tempplayer As Long
    tempplayer = CLng(Mid$(strData(i), 11, 1))
    sTotalDjinn(tempplayer) = Mid(strData(i), 12, Len(strData(i)))
    iTotalDjinn(tempplayer) = CInt(sTotalDjinn(tempplayer))
End If


If Left$(strData(i), 6) = "CURPSY" Then
    sCurPsy = Mid(strData(i), 7, Len(strData(i)))
    iCurPsy = CInt(sCurPsy)
End If

If Left$(strData(i), 7) = "PSYNAME" Then
    strPsyName(iCurPsy) = Mid(strData(i), 8, Len(strData(i)))
    Psynergy(iCurPsy).Name = strPsyName(iCurPsy)
End If

If Left$(strData(i), 6) = "PSYDMG" Then
    strPsyDamage(iCurPsy) = Mid(strData(i), 7, Len(strData(i)))
    Psynergy(iCurPsy).Damage = CLng(strPsyDamage(iCurPsy))
End If

If Left$(strData(i), 7) = "PSYTYPE" Then
    strPsyType(iCurPsy) = Mid(strData(i), 8, Len(strData(i)))
    Psynergy(iCurPsy).Type = strPsyType(iCurPsy)
'    If Mid$(strPsyType(iCurPsy), 0, 3) = "DAM" Then strPsyType(iCurPsy) = "DAMAGE"
End If

If Left$(strData(i), 5) = "PSYPP" Then
    strPsyPP(iCurPsy) = Mid(strData(i), 6, Len(strData(i)))
    Psynergy(iCurPsy).PP = CLng(strPsyPP(iCurPsy))
End If

If Left$(strData(i), 8) = "PSYDJINN" Then
    strPsyDjinn(iCurPsy) = Mid(strData(i), 9, Len(strData(i)))
    Psynergy(iCurPsy).Djinn = CLng(strPsyDjinn(iCurPsy))
End If

If Left$(strData(i), 7) = "PSYDESC" Then
    strPsyDesc(iCurPsy) = Mid(strData(i), 8, Len(strData(i)))
    Psynergy(iCurPsy).Desc = strPsyDesc(iCurPsy)
End If

If Left$(strData(i), 9) = "PSYPLAYER" Then
    Psynergy(iCurPsy).Character = CLng(Mid(strData(i), 10, Len(strData(i))))
End If


If Left$(strData(i), 6) = "CURSUM" Then
    sCurSum = Mid(strData(i), 7, Len(strData(i)))
    iCurSum = CInt(sCurSum)
End If

If Left$(strData(i), 7) = "SUMNAME" Then
    strSumName(iCurSum) = Mid(strData(i), 8, Len(strData(i)))
    Summon(iCurSum).Name = strSumName(iCurSum)
End If

If Left$(strData(i), 8) = "SUMDJINN" Then
    strSumDjinn(iCurSum) = Mid(strData(i), 9, Len(strData(i)))
    Summon(iCurSum).Level = strSumDjinn(iCurSum)
End If

If Left$(strData(i), 7) = "SUMDESC" Then
    strSumDesc(iCurSum) = Mid(strData(i), 8, Len(strData(i)))
    Summon(iCurSum).Desc = strSumDesc(iCurSum)
End If

If Left$(strData(i), 7) = "SUMCHAR" Then
    strSumDesc(iCurSum) = Mid(strData(i), 8, Len(strData(i)))
    Summon(iCurSum).Character = CLng(strSumDesc(iCurSum))
End If

If Left$(strData(i), 7) = "CUSTNUM" Then
    iCurCust = CInt(Mid(strData(i), 8, Len(strData(i))))
End If
If Left$(strData(i), 8) = "CUSTNAME" Then
    CustomChar(iCurCust).Name = Mid(strData(i), 9, Len(strData(i)))
End If
If Left$(strData(i), 11) = "CUSTPICTURE" Then
    CustomChar(iCurCust).Picture = Mid(strData(i), 12, Len(strData(i)))
End If
If Left$(strData(i), 6) = "CUSTHP" Then
    CustomChar(iCurCust).BaseHP = CInt(Mid(strData(i), 7, Len(strData(i))))
End If
If Left$(strData(i), 6) = "CUSTAP" Then
    CustomChar(iCurCust).BaseAP = CInt(Mid(strData(i), 7, Len(strData(i))))
End If
If Left$(strData(i), 11) = "CUSTDEFENSE" Then
    CustomChar(iCurCust).BaseDefense = CInt(Mid(strData(i), 12, Len(strData(i))))
End If
If Left$(strData(i), 7) = "CUSTRES" Then
    CustomChar(iCurCust).BaseRes = CInt(Mid(strData(i), 8, Len(strData(i))))
End If
If Left$(strData(i), 9) = "CUSTPOWER" Then
    CustomChar(iCurCust).BasePower = CInt(Mid(strData(i), 10, Len(strData(i))))
End If
If Left$(strData(i), 6) = "CUSTPP" Then
    CustomChar(iCurCust).BasePP = CInt(Mid(strData(i), 7, Len(strData(i))))
End If
If Left$(strData(i), 12) = "CUSTSTRENGTH" Then
    CustomChar(iCurCust).Strength = Mid(strData(i), 13, Len(strData(i)))
End If
If Left$(strData(i), 12) = "CUSTWEAKNESS" Then
    CustomChar(iCurCust).Weakness = Mid(strData(i), 13, Len(strData(i)))
End If
If Left$(strData(i), 8) = "CUSTLUCK" Then
    CustomChar(iCurCust).BaseLuck = CInt(Mid(strData(i), 9, Len(strData(i))))
End If
If Left$(strData(i), 8) = "CUSTTYPE" Then
    CustomChar(iCurCust).Type = Mid(strData(i), 9, Len(strData(i)))
End If
If Left$(strData(i), 8) = "CUSTDESC" Then
    CustomChar(iCurCust).Description = Mid(strData(i), 9, Len(strData(i)))
End If
If Left$(strData(i), 8) = "CUSTUSER" Then
    CustomChar(iCurCust).Users = Mid(strData(i), 9, Len(strData(i)))
    Dim strCustArray
    strCustArray = Split(CustomChar(iCurCust).Users, "@", -1, vbTextCompare)
    For q = 0 To UBound(strCustArray)
        If strMyUserName = strCustArray(q) Or strCustArray(q) = "ANY" Then
            cmbCharPic.AddItem CustomChar(iCurCust).Name
        End If
    Next 'q
End If

If Left$(strData(i), 6) = "RATING" Then
    strRating = Mid(strData(i), 7, Len(strData(i)))
    lblRating.Caption = strRating
End If
If Left$(strData(i), 4) = "WINS" Then
    strWins = Mid(strData(i), 5, Len(strData(i)))
    lblWins.Caption = strWins
End If
If Left$(strData(i), 4) = "LOSS" Then
    strLoss = Mid(strData(i), 5, Len(strData(i)))
    lblLosses.Caption = strLoss
End If
If Left$(strData(i), 4) = "DISC" Then
    strDisc = Mid(strData(i), 5, Len(strData(i)))
    lblDisc.Caption = strDisc
End If
If Left$(strData(i), 5) = "COINS" Then
    strCoins = Mid(strData(i), 6, Len(strData(i)))
End If
'If Left$(strdata(i), 5) = "DJINN" Then
'    strDjinn = Mid(strdata(i), 6, Len(strdata(i)))
'End If
If Left$(strData(i), 4) = "CHAR" Then
    strChar(1) = Mid(strData(i), 5, Len(strData(i)))
    CharName(1) = strChar(1)
End If
If Left$(strData(i), 5) = "2CHAR" Then
    strChar(2) = Mid(strData(i), 6, Len(strData(i)))
    CharName(2) = strChar(2)
End If

If Left$(strData(i), 3) = "WPN" Then
    strWeapon(1) = Mid(strData(i), 4, Len(strData(i)))
    intWeapon(1) = CInt(strWeapon(1))
End If
If Left$(strData(i), 4) = "2WPN" Then
    strWeapon(2) = Mid(strData(i), 5, Len(strData(i)))
    intWeapon(2) = CInt(strWeapon(2))
End If


If Left$(strData(i), 5) = "KARMA" Then
    Dim strKarma As String
    strKarma = Mid(strData(i), 6, Len(strData(i)))
    If strKarma = "" Then strKarma = "0"
    intKarma = CInt(strKarma)
End If

If Left$(strData(i), 3) = "LVL" Then
    strLvl = Mid(strData(i), 4, Len(strData(i)))
    lblLevel.Caption = strLvl
     
    LoggedIn = True
     
End If

    If Left$(strData(i), 10) = "ITEMCONFIRM" Then
        lblStatus(1).Caption = "Item updated!"
    End If
     
If Left$(strData(i), 3) = "TOS" Then
    txtTOS.Text = txtTOS.Text & vbNewLine & Mid(strData(i), 4, Len(strData(i)))
    lblStatus(1).Caption = "Connected to server."
End If

If Left$(strData(i), 11) = "SINGLECOINS" Then
    MsgBox "Your coins were updated successfully!"
End If

'If Left$(strdata(i), 7) = "NUMUSER" Then
'    ServerNumber = CInt(Mid(strdata(i), 8, Len(strdata(i))))
'End If

If Left$(strData(i), 11) = "HIGHSCOREDS" Then
    intDjinnSaveHighScore = CInt(Mid(strData(i), 12, Len(strData(i))))
End If
If Left$(strData(i), 12) = "HIGHPLAYERDS" Then
    strDjinnSavePlayer = Mid(strData(i), 13, Len(strData(i)))
End If

If Left$(strData(i), 4) = "FULL" Then
    lblStatus(1).Caption = "All server sockets are full at this time.  Please try again later."
    MsgBox "All server sockets are full at this time.  Please try again later."
End If

If Left$(strData(i), 6) = "BADPIN" Then
    MsgBox "You were unable to log-in because you have been banned from the game."
End If

If Left$(strData(i), 7) = "MODNAME" Then
    Dim strAddMod As String
    strAddMod = Mid$(strData(i), 8, Len(strData(i)))
    For q = 1 To 15
        If strModName(q) = "" Then
            strModName(q) = strAddMod
            Exit For
        End If
    Next 'q
End If

If Left$(strData(i), 4) = "DONE" Then
    cmdLogin.Enabled = False
    cmdLogin.Caption = "Logged In"
    txtUserName.Enabled = False
    txtPassword.Enabled = False
    cmdProceed.Enabled = True
    lblStatus(1).Caption = "Logged In To The Server"
End If

If Left$(strData(i), 12) = "CHANGEPWGOOD" Then
    MsgBox "Password changed successfully!"
End If
If Left$(strData(i), 11) = "CHANGEPWBAD" Then
    MsgBox "Password not changed.  Make sure that you have entered your old password correctly."
End If


'If Left$(strdata(i), 3) = "LST" Then    'if we got file list packet
'    lblStatus(1) = "Downloading game data. This may take a minute." 'Update Status
'    FileNameList = Split(Mid$(strdata(i), 4), ",", -1, vbTextCompare) 'Create file name array
'    intCurFile = 0
'    DownloadNextFile  'goto the download sub
'End If

'If Left$(strdata(i), 4) = "FILE" Then
'    If GettingFile = False Then  'if not currently getting file
'        GettingFile = True 'now we are getting file
'        File = Mid$(strdata(i), 5, Len(strdata(i))) 'replace old variable value with data
'    Else    'otherwise we are getting file
'        File = File & Mid$(strdata(i), 5, Len(strdata(i)))  'add data to variable
'    End If
'    If Right$(strdata(i), 1) = "!" Then    'if the end of file marker was found
'        GettingFile = False 'we are no longer reciving a file
'        Call MakeFile(File, FileNameList(intCurFile))   'make the file
'        intCurFile = intCurFile + 1
'        Call DownloadNextFile
'    Else
'        Debug.Print strdata(i)
'    End If
'End If


Next 'i

Debug.Print strdatao


End Sub


Public Function Decode(sData As String) As String
    Dim sTemp As String, sTemp1 As String


    For II% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, II%, 1)
        lT = Asc(sTemp$) \ 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next II%
    Decode$ = sTemp1$
End Function
Sub TextShow()
If bNewChar = True Then
    cmdLogin.Caption = "Create New"
    cmdProceed.Visible = False
    For i = 0 To lblStats.UBound
        lblStats(i).Visible = False
    Next 'i
    lblWins.Visible = False
    lblLosses.Visible = False
    lblDisc.Visible = False
    lblRating.Visible = False
    lblLevel.Visible = False
    lblChar(0).Visible = True
    lblChar(1).Visible = True
    lblChar(2).Visible = True
    lblChar(3).Visible = True
    lblChar(4).Visible = True
    lblChar(5).Visible = True
    cmbCharPic.Visible = True
    cmbChar2.Visible = True
    chkSaveLogin.Visible = False
    cmdChangePW.Visible = False
Else
    cmdLogin.Caption = "Log In"
    cmdProceed.Visible = True
    For i = 0 To lblStats.UBound
        lblStats(i).Visible = True
    Next 'i
    txtUserName.Text = ""
    txtPassword.Text = ""
    lblWins.Visible = True
    lblLosses.Visible = True
    lblDisc.Visible = True
    lblRating.Visible = True
    lblLevel.Visible = True
    lblChar(0).Visible = False
    lblChar(1).Visible = False
    lblChar(2).Visible = False
    lblChar(3).Visible = False
    lblChar(4).Visible = False
    lblChar(5).Visible = False
    cmbCharPic.Visible = False
    cmbChar2.Visible = False
    chkSaveLogin.Visible = True
    cmdChangePW.Visible = True
End If
End Sub

Private Sub DownloadNextFile()
On Error Resume Next

If intCurFile < UBound(FileNameList) + 1 Then 'if there are more files

    Dim bExist As Boolean
    bExist = False
    For i = 0 To lstFile.ListCount
        'Don't donwload if it is not a maze file and you already have it
        If FileNameList(intCurFile) = lstFile.List(i) And Right$(FileNameList(intCurFile), 4) <> "omaz" Then
            bExist = True
            Exit For
        End If
    Next 'i
    If bExist = True Then
        Call FileExists ' if the file exists goto the file exists handler
        Exit Sub
    Else
        FileTransfer.SendData "RQSTFL" & FileNameList(intCurFile) & vbCrLf
    End If
    
    Exit Sub 'exit sub until File has been recived

Else 'Maximum number of files downloaded

    'If strIkillkennyPW = "z1x2" Then
    '    Call SendAdminData
    'Else
        'User.SendData "RQSTFL" & FileNameList(intCurFile) & vbCrLf
'        FileTransfer.SendData "CLOSE" & vbCrLf
'        DoEvents
        FileTransfer.Close
'        Exit Sub
    'End If

End If

End Sub
Sub FileExists()
intCurFile = intCurFile + 1 'increase the counter
Call DownloadNextFile
Exit Sub

End Sub
Public Sub MakeFile(ByVal File As String)
Dim FileName As String, filedata As String
On Error Resume Next
'If Left$(File, 4) = "FILE" Then 'check for propper header
    File = Right$(File, Len(File) - 4)  'remove header
    FileName = Left$(File, InStr(File, "@") - 1) 'get the filename
    FileName = App.Path & "\files\" & FileName
    filedata = Mid$(File, InStr(File, "@") + 1, InStr(File, "!") - 1 - InStr(File, "@")) 'get the file data
'    Debug.Assert Len(filename) + Len(filedata) + 2 = Len(File)
    
    '=================================
'    With CommonDialog1
'        .filename = filename
'        .DialogTitle = "Save"
'        .ShowSave
'        filename = .filename
'    End With
    Open FileName For Binary Access Write As #1
    Put #1, , filedata
    Close #1
    '=================================
    b64.DecodeFile FileName, FileName
Close #1
'End If
End Sub
Sub SendAdminData()
On Error Resume Next
Dim strInput As String
strInput = InputBox("Enter desired file.  Enter nothing to break the loop.")
If strInput = "" Then
    FileTransfer.SendData "CLOSE" & vbCrLf
    DoEvents
    FileTransfer.Close
    Exit Sub
End If
Dim strSplit
strSplit = Split(strInput, "@", -1, vbTextCompare)
For i = 0 To UBound(strSplit)
    FileTransfer.SendData "RQSTFL" & strSplit(i) & vbCrLf
    DoEvents
Next 'i
End Sub
