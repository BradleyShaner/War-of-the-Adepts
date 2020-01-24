VERSION 5.00
Begin VB.Form frmBattle2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "War of the Adepts - Battle"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8250
   ControlBox      =   0   'False
   Icon            =   "frmBattle2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtChatMsg 
      BackColor       =   &H00000000&
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
      MaxLength       =   150
      TabIndex        =   107
      Top             =   6120
      Width           =   8055
   End
   Begin VB.TextBox txtChat 
      BackColor       =   &H00000000&
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
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   106
      Text            =   "frmBattle2.frx":08CA
      Top             =   5400
      Width           =   8055
   End
   Begin VB.Frame framStatus 
      BackColor       =   &H00886000&
      Caption         =   "Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   1200
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lblSArmor 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
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
         Height          =   240
         Index           =   2
         Left            =   2880
         TabIndex        =   192
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblSArmor 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
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
         Height          =   240
         Index           =   1
         Left            =   2880
         TabIndex        =   191
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblSArmor 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
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
         Height          =   240
         Index           =   0
         Left            =   2880
         TabIndex        =   190
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblSWeapon 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
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
         Height          =   240
         Left            =   3120
         TabIndex        =   189
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Armor:"
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
         Height          =   240
         Index           =   16
         Left            =   2160
         TabIndex        =   188
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Armor:"
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
         Height          =   240
         Index           =   15
         Left            =   2160
         TabIndex        =   187
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Armor:"
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
         Height          =   240
         Index           =   14
         Left            =   2160
         TabIndex        =   186
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weapon:"
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
         Height          =   240
         Index           =   13
         Left            =   2160
         TabIndex        =   185
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblSNResist 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   4560
         TabIndex        =   63
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblSFResist 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   3600
         TabIndex        =   62
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblSWResist 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   2640
         TabIndex        =   61
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblSEResist 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   1680
         TabIndex        =   60
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblSNPower 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   4560
         TabIndex        =   59
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblSFPower 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   3600
         TabIndex        =   58
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblSWPower 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   2640
         TabIndex        =   57
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblSEPower 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   1680
         TabIndex        =   56
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblSNLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   4560
         TabIndex        =   55
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblSFLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   3600
         TabIndex        =   54
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblSWLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   2640
         TabIndex        =   53
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblSELevel 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   1680
         TabIndex        =   52
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblSNDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   4560
         TabIndex        =   51
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblSFDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   3600
         TabIndex        =   50
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblSWDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   2640
         TabIndex        =   49
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblSEDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   240
         Left            =   1680
         TabIndex        =   48
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblSLuck 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5040
         TabIndex        =   47
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblSAgility 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5160
         TabIndex        =   46
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblSDefense 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5280
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblSAttack 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5040
         TabIndex        =   44
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblSPP 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   1320
         TabIndex        =   43
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblSHP 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   1320
         TabIndex        =   42
         Top             =   600
         Width           =   1215
      End
      Begin VB.Image imgSPic 
         Height          =   960
         Left            =   5280
         Picture         =   "frmBattle2.frx":08E5
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label lblSStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
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
         Height          =   240
         Left            =   1560
         TabIndex        =   40
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblSLvl 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Left            =   1200
         TabIndex        =   39
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblSExp 
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
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
         Height          =   240
         Left            =   840
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
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
         Height          =   240
         Index           =   12
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lv."
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
         Height          =   240
         Index           =   11
         Left            =   840
         TabIndex        =   36
         Top             =   120
         Width           =   285
      End
      Begin VB.Label lblSSwitchR 
         BackStyle       =   0  'Transparent
         Caption         =   "Switch Right"
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
         Height          =   240
         Left            =   5280
         TabIndex        =   35
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lblSSwitchL 
         BackStyle       =   0  'Transparent
         Caption         =   "Switch Left"
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
         Height          =   240
         Left            =   4080
         TabIndex        =   34
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   60
      End
      Begin VB.Image imgSBall 
         Height          =   210
         Index           =   3
         Left            =   2640
         Picture         =   "frmBattle2.frx":0FF5
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgSBall 
         Height          =   210
         Index           =   2
         Left            =   4560
         Picture         =   "frmBattle2.frx":1337
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgSBall 
         Height          =   210
         Index           =   1
         Left            =   3600
         Picture         =   "frmBattle2.frx":1679
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgSBall 
         Height          =   210
         Index           =   0
         Left            =   1680
         Picture         =   "frmBattle2.frx":19BB
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resist:"
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
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   32
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Power:"
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
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   31
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lv:"
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
         Height          =   240
         Index           =   7
         Left            =   240
         TabIndex        =   30
         Top             =   2160
         Width           =   285
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Djinn:"
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
         Height          =   240
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   1920
         Width           =   555
      End
      Begin VB.Image imgSFace 
         Height          =   480
         Left            =   240
         Picture         =   "frmBattle2.frx":1CFD
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblSClass 
         BackStyle       =   0  'Transparent
         Caption         =   "Lord"
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
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luck:"
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
         Height          =   240
         Index           =   5
         Left            =   4320
         TabIndex        =   27
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agility:"
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
         Height          =   240
         Index           =   4
         Left            =   4320
         TabIndex        =   26
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
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
         Height          =   240
         Index           =   3
         Left            =   4320
         TabIndex        =   25
         Top             =   840
         Width           =   825
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attack:"
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
         Height          =   240
         Index           =   2
         Left            =   4320
         TabIndex        =   24
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PP:"
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
         Height          =   240
         Index           =   1
         Left            =   840
         TabIndex        =   23
         Top             =   840
         Width           =   330
      End
      Begin VB.Label lblSGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
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
         Height          =   240
         Index           =   0
         Left            =   840
         TabIndex        =   22
         Top             =   600
         Width           =   330
      End
      Begin VB.Label lblSChar 
         BackStyle       =   0  'Transparent
         Caption         =   "Felix"
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
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame framItems 
      BackColor       =   &H00886000&
      Caption         =   "Choose Your Item"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   1200
      TabIndex        =   178
      Top             =   1440
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restores 20 HP."
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
         Height          =   240
         Index           =   2
         Left            =   1680
         MouseIcon       =   "frmBattle2.frx":1FDB
         TabIndex        =   184
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restores 20 HP."
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
         Height          =   240
         Index           =   1
         Left            =   1680
         MouseIcon       =   "frmBattle2.frx":28A5
         TabIndex        =   183
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restores 20 HP."
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
         Height          =   240
         Index           =   0
         Left            =   1680
         MouseIcon       =   "frmBattle2.frx":316F
         TabIndex        =   182
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Herb"
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
         Height          =   240
         Index           =   2
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":3A39
         MousePointer    =   99  'Custom
         TabIndex        =   181
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Herb"
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
         Height          =   240
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":4303
         MousePointer    =   99  'Custom
         TabIndex        =   180
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Herb"
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
         Height          =   240
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":4BCD
         MousePointer    =   99  'Custom
         TabIndex        =   179
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgItemIcon 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "frmBattle2.frx":5497
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgItemIcon 
         Height          =   240
         Index           =   1
         Left            =   240
         Picture         =   "frmBattle2.frx":55EA
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgItemIcon 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmBattle2.frx":573D
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame framSummon 
      BackColor       =   &H00886000&
      Caption         =   "Choose Your Summon"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   1200
      TabIndex        =   147
      Top             =   1440
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   9
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":5890
         TabIndex        =   177
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   8
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":615A
         TabIndex        =   176
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   7
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":6A24
         TabIndex        =   175
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   6
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":72EE
         TabIndex        =   174
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   5
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":7BB8
         TabIndex        =   173
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   4
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":8482
         TabIndex        =   172
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   3
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":8D4C
         TabIndex        =   171
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   2
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":9616
         TabIndex        =   170
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   1
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":9EE0
         TabIndex        =   169
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblSummonDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   240
         Index           =   0
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":A7AA
         TabIndex        =   168
         Top             =   360
         Width           =   255
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   9
         Left            =   5040
         Picture         =   "frmBattle2.frx":B074
         Top             =   2565
         Width           =   90
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   8
         Left            =   5040
         Picture         =   "frmBattle2.frx":B3B6
         Top             =   2325
         Width           =   90
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   7
         Left            =   5040
         Picture         =   "frmBattle2.frx":B6F8
         Top             =   2085
         Width           =   90
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   6
         Left            =   5040
         Picture         =   "frmBattle2.frx":BA3A
         Top             =   1845
         Width           =   90
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   5
         Left            =   5040
         Picture         =   "frmBattle2.frx":BD7C
         Top             =   1605
         Width           =   90
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   4
         Left            =   5040
         Picture         =   "frmBattle2.frx":C0BE
         Top             =   1365
         Width           =   90
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   3
         Left            =   5040
         Picture         =   "frmBattle2.frx":C400
         Top             =   1125
         Width           =   90
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   2
         Left            =   5040
         Picture         =   "frmBattle2.frx":C742
         Top             =   885
         Width           =   90
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   1
         Left            =   5040
         Picture         =   "frmBattle2.frx":CA84
         Top             =   645
         Width           =   90
      End
      Begin VB.Image imgSummonType 
         Height          =   90
         Index           =   0
         Left            =   5040
         Picture         =   "frmBattle2.frx":CDC6
         Top             =   405
         Width           =   90
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   9
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":D108
         TabIndex        =   167
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   8
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":D9D2
         TabIndex        =   166
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   7
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":E29C
         TabIndex        =   165
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   6
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":EB66
         TabIndex        =   164
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   5
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":F430
         TabIndex        =   163
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":FCFA
         TabIndex        =   162
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   3
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":105C4
         TabIndex        =   161
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   2
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":10E8E
         TabIndex        =   160
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":11758
         TabIndex        =   159
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblSummonDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "The elemental power of earth."
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
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmBattle2.frx":12022
         TabIndex        =   158
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   9
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":128EC
         MousePointer    =   99  'Custom
         TabIndex        =   157
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   8
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":131B6
         MousePointer    =   99  'Custom
         TabIndex        =   156
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   7
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":13A80
         MousePointer    =   99  'Custom
         TabIndex        =   155
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   6
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":1434A
         MousePointer    =   99  'Custom
         TabIndex        =   154
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   5
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":14C14
         MousePointer    =   99  'Custom
         TabIndex        =   153
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   4
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":154DE
         MousePointer    =   99  'Custom
         TabIndex        =   152
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   3
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":15DA8
         MousePointer    =   99  'Custom
         TabIndex        =   151
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   2
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":16672
         MousePointer    =   99  'Custom
         TabIndex        =   150
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":16F3C
         MousePointer    =   99  'Custom
         TabIndex        =   149
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblSummon 
         BackStyle       =   0  'Transparent
         Caption         =   "Venus"
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
         Height          =   240
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmBattle2.frx":17806
         MousePointer    =   99  'Custom
         TabIndex        =   148
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   9
         Left            =   240
         Picture         =   "frmBattle2.frx":180D0
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   8
         Left            =   240
         Picture         =   "frmBattle2.frx":1845E
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmBattle2.frx":187EC
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   6
         Left            =   240
         Picture         =   "frmBattle2.frx":18B7A
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   5
         Left            =   240
         Picture         =   "frmBattle2.frx":18F08
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   4
         Left            =   240
         Picture         =   "frmBattle2.frx":19296
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   3
         Left            =   240
         Picture         =   "frmBattle2.frx":19624
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "frmBattle2.frx":199B2
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   1
         Left            =   240
         Picture         =   "frmBattle2.frx":19D40
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgSummonIcon 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmBattle2.frx":1A0CE
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame framDjinn 
      BackColor       =   &H00886000&
      Caption         =   "Choose Your Djinni"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   1200
      TabIndex        =   108
      Top             =   1440
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Image imgStatDown 
         Height          =   150
         Left            =   5160
         Picture         =   "frmBattle2.frx":1A45C
         Top             =   2280
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image imgStatBoost 
         Height          =   150
         Left            =   4920
         Picture         =   "frmBattle2.frx":1A595
         Top             =   2280
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image imgDjinnStat 
         Height          =   150
         Index           =   5
         Left            =   5880
         Picture         =   "frmBattle2.frx":1A6CE
         Top             =   1800
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image imgDjinnStat 
         Height          =   150
         Index           =   4
         Left            =   5880
         Picture         =   "frmBattle2.frx":1A807
         Top             =   1560
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image imgDjinnStat 
         Height          =   150
         Index           =   3
         Left            =   5880
         Picture         =   "frmBattle2.frx":1A940
         Top             =   1320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image imgDjinnStat 
         Height          =   150
         Index           =   2
         Left            =   5880
         Picture         =   "frmBattle2.frx":1AA79
         Top             =   1080
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image imgDjinnStat 
         Height          =   150
         Index           =   1
         Left            =   5880
         Picture         =   "frmBattle2.frx":1ABB2
         Top             =   840
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image imgDjinnStat 
         Height          =   150
         Index           =   0
         Left            =   5880
         Picture         =   "frmBattle2.frx":1ACEB
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblDNLuck 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   6120
         MouseIcon       =   "frmBattle2.frx":1AE24
         MousePointer    =   99  'Custom
         TabIndex        =   146
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblDNAgility 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   6120
         MouseIcon       =   "frmBattle2.frx":1B6EE
         MousePointer    =   99  'Custom
         TabIndex        =   145
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblDNDefense 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   6120
         MouseIcon       =   "frmBattle2.frx":1BFB8
         MousePointer    =   99  'Custom
         TabIndex        =   144
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblDNAttack 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   6120
         MouseIcon       =   "frmBattle2.frx":1C882
         MousePointer    =   99  'Custom
         TabIndex        =   143
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblDNPP 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   6120
         MouseIcon       =   "frmBattle2.frx":1D14C
         MousePointer    =   99  'Custom
         TabIndex        =   142
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblDNHP 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   6120
         MouseIcon       =   "frmBattle2.frx":1DA16
         MousePointer    =   99  'Custom
         TabIndex        =   141
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblDOLuck 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5160
         MouseIcon       =   "frmBattle2.frx":1E2E0
         MousePointer    =   99  'Custom
         TabIndex        =   140
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblDOAgility 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":1EBAA
         MousePointer    =   99  'Custom
         TabIndex        =   139
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblDODefense 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5520
         MouseIcon       =   "frmBattle2.frx":1F474
         MousePointer    =   99  'Custom
         TabIndex        =   138
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblDOAttack 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5280
         MouseIcon       =   "frmBattle2.frx":1FD3E
         MousePointer    =   99  'Custom
         TabIndex        =   137
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblDOPP 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5040
         MouseIcon       =   "frmBattle2.frx":20608
         MousePointer    =   99  'Custom
         TabIndex        =   136
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblDOHP 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   240
         Left            =   5040
         MouseIcon       =   "frmBattle2.frx":20ED2
         MousePointer    =   99  'Custom
         TabIndex        =   135
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblDjinnGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luck:"
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
         Height          =   210
         Index           =   5
         Left            =   4680
         MouseIcon       =   "frmBattle2.frx":2179C
         MousePointer    =   99  'Custom
         TabIndex        =   134
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label lblDjinnGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agility:"
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
         Height          =   210
         Index           =   4
         Left            =   4680
         MouseIcon       =   "frmBattle2.frx":22066
         MousePointer    =   99  'Custom
         TabIndex        =   133
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label lblDjinnGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
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
         Height          =   210
         Index           =   3
         Left            =   4680
         MouseIcon       =   "frmBattle2.frx":22930
         MousePointer    =   99  'Custom
         TabIndex        =   132
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblDjinnGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attack:"
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
         Height          =   210
         Index           =   2
         Left            =   4680
         MouseIcon       =   "frmBattle2.frx":231FA
         MousePointer    =   99  'Custom
         TabIndex        =   131
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lblDjinnGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PP:"
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
         Height          =   210
         Index           =   1
         Left            =   4680
         MouseIcon       =   "frmBattle2.frx":23AC4
         MousePointer    =   99  'Custom
         TabIndex        =   130
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblDjinnGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
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
         Height          =   210
         Index           =   0
         Left            =   4680
         MouseIcon       =   "frmBattle2.frx":2438E
         MousePointer    =   99  'Custom
         TabIndex        =   129
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblDNClass 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lord"
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
         Height          =   240
         Left            =   5520
         MouseIcon       =   "frmBattle2.frx":24C58
         MousePointer    =   99  'Custom
         TabIndex        =   128
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDOClass 
         BackStyle       =   0  'Transparent
         Caption         =   "Lord"
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
         Height          =   240
         Left            =   4680
         MouseIcon       =   "frmBattle2.frx":25522
         MousePointer    =   99  'Custom
         TabIndex        =   127
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDjinnDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Strike a blow that can cleave stone."
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
         Height          =   240
         Index           =   8
         Left            =   1440
         MouseIcon       =   "frmBattle2.frx":25DEC
         TabIndex        =   126
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label lblDjinnDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Strike a blow that can cleave stone."
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
         Height          =   240
         Index           =   7
         Left            =   1440
         MouseIcon       =   "frmBattle2.frx":266B6
         TabIndex        =   125
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label lblDjinnDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Strike a blow that can cleave stone."
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
         Height          =   240
         Index           =   6
         Left            =   1440
         MouseIcon       =   "frmBattle2.frx":26F80
         TabIndex        =   124
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lblDjinnDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Strike a blow that can cleave stone."
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
         Height          =   240
         Index           =   5
         Left            =   1440
         MouseIcon       =   "frmBattle2.frx":2784A
         TabIndex        =   123
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label lblDjinnDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Strike a blow that can cleave stone."
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
         Height          =   240
         Index           =   4
         Left            =   1440
         MouseIcon       =   "frmBattle2.frx":28114
         TabIndex        =   122
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label lblDjinnDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Strike a blow that can cleave stone."
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
         Height          =   240
         Index           =   3
         Left            =   1440
         MouseIcon       =   "frmBattle2.frx":289DE
         TabIndex        =   121
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblDjinnDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Strike a blow that can cleave stone."
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
         Height          =   240
         Index           =   2
         Left            =   1440
         MouseIcon       =   "frmBattle2.frx":292A8
         TabIndex        =   120
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lblDjinnDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Strike a blow that can cleave stone."
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
         Height          =   240
         Index           =   1
         Left            =   1440
         MouseIcon       =   "frmBattle2.frx":29B72
         TabIndex        =   119
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblDjinnDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Strike a blow that can cleave stone."
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
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmBattle2.frx":2A43C
         TabIndex        =   118
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "Flint"
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
         Height          =   240
         Index           =   8
         Left            =   480
         MouseIcon       =   "frmBattle2.frx":2AD06
         MousePointer    =   99  'Custom
         TabIndex        =   117
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "Flint"
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
         Height          =   240
         Index           =   7
         Left            =   480
         MouseIcon       =   "frmBattle2.frx":2B5D0
         MousePointer    =   99  'Custom
         TabIndex        =   116
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "Flint"
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
         Height          =   240
         Index           =   6
         Left            =   480
         MouseIcon       =   "frmBattle2.frx":2BE9A
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "Flint"
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
         Height          =   240
         Index           =   5
         Left            =   480
         MouseIcon       =   "frmBattle2.frx":2C764
         MousePointer    =   99  'Custom
         TabIndex        =   114
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "Flint"
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
         Height          =   240
         Index           =   4
         Left            =   480
         MouseIcon       =   "frmBattle2.frx":2D02E
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "Flint"
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
         Height          =   240
         Index           =   3
         Left            =   480
         MouseIcon       =   "frmBattle2.frx":2D8F8
         MousePointer    =   99  'Custom
         TabIndex        =   112
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "Flint"
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
         Height          =   240
         Index           =   2
         Left            =   480
         MouseIcon       =   "frmBattle2.frx":2E1C2
         MousePointer    =   99  'Custom
         TabIndex        =   111
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "Flint"
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
         Height          =   240
         Index           =   1
         Left            =   480
         MouseIcon       =   "frmBattle2.frx":2EA8C
         MousePointer    =   99  'Custom
         TabIndex        =   110
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblDjinn 
         BackStyle       =   0  'Transparent
         Caption         =   "Flint"
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
         Height          =   240
         Index           =   0
         Left            =   480
         MouseIcon       =   "frmBattle2.frx":2F356
         MousePointer    =   99  'Custom
         TabIndex        =   109
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgDjinnType 
         Height          =   90
         Index           =   8
         Left            =   240
         Picture         =   "frmBattle2.frx":2FC20
         Top             =   2325
         Width           =   90
      End
      Begin VB.Image imgDjinnType 
         Height          =   90
         Index           =   7
         Left            =   240
         Picture         =   "frmBattle2.frx":2FF62
         Top             =   2085
         Width           =   90
      End
      Begin VB.Image imgDjinnType 
         Height          =   90
         Index           =   6
         Left            =   240
         Picture         =   "frmBattle2.frx":302A4
         Top             =   1845
         Width           =   90
      End
      Begin VB.Image imgDjinnType 
         Height          =   90
         Index           =   5
         Left            =   240
         Picture         =   "frmBattle2.frx":305E6
         Top             =   1605
         Width           =   90
      End
      Begin VB.Image imgDjinnType 
         Height          =   90
         Index           =   4
         Left            =   240
         Picture         =   "frmBattle2.frx":30928
         Top             =   1365
         Width           =   90
      End
      Begin VB.Image imgDjinnType 
         Height          =   90
         Index           =   3
         Left            =   240
         Picture         =   "frmBattle2.frx":30C6A
         Top             =   1125
         Width           =   90
      End
      Begin VB.Image imgDjinnType 
         Height          =   90
         Index           =   2
         Left            =   240
         Picture         =   "frmBattle2.frx":30FAC
         Top             =   885
         Width           =   90
      End
      Begin VB.Image imgDjinnType 
         Height          =   90
         Index           =   1
         Left            =   240
         Picture         =   "frmBattle2.frx":312EE
         Top             =   645
         Width           =   90
      End
      Begin VB.Image imgDjinnType 
         Height          =   90
         Index           =   0
         Left            =   240
         Picture         =   "frmBattle2.frx":31630
         Top             =   405
         Width           =   90
      End
   End
   Begin VB.Frame framPsynergy 
      BackColor       =   &H00886000&
      Caption         =   "Choose Your Psynergy"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   1200
      TabIndex        =   64
      Top             =   1440
      Visible         =   0   'False
      Width           =   6615
      Begin VB.VScrollBar psyScroll 
         Height          =   2415
         LargeChange     =   5
         Left            =   6240
         Max             =   5
         TabIndex        =   95
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   9
         Left            =   3120
         Picture         =   "frmBattle2.frx":31972
         Top             =   2445
         Width           =   360
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   8
         Left            =   3120
         Picture         =   "frmBattle2.frx":31A8E
         Top             =   2205
         Width           =   360
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   7
         Left            =   3120
         Picture         =   "frmBattle2.frx":31BAA
         Top             =   1965
         Width           =   360
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   6
         Left            =   3120
         Picture         =   "frmBattle2.frx":31CC6
         Top             =   1725
         Width           =   360
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   5
         Left            =   3120
         Picture         =   "frmBattle2.frx":31DE2
         Top             =   1485
         Width           =   360
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   4
         Left            =   3120
         Picture         =   "frmBattle2.frx":31EFE
         Top             =   1245
         Width           =   360
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   3
         Left            =   3120
         Picture         =   "frmBattle2.frx":3201A
         Top             =   1005
         Width           =   360
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   2
         Left            =   3120
         Picture         =   "frmBattle2.frx":32136
         Top             =   765
         Width           =   360
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   1
         Left            =   3120
         Picture         =   "frmBattle2.frx":32252
         Top             =   525
         Width           =   360
      End
      Begin VB.Image imgPsyRange 
         Height          =   120
         Index           =   0
         Left            =   3120
         Picture         =   "frmBattle2.frx":3236E
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   9
         Left            =   3600
         TabIndex        =   105
         Top             =   2400
         Width           =   2490
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   8
         Left            =   3600
         TabIndex        =   104
         Top             =   2160
         Width           =   2490
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   7
         Left            =   3600
         TabIndex        =   103
         Top             =   1920
         Width           =   2490
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   6
         Left            =   3600
         TabIndex        =   102
         Top             =   1680
         Width           =   2490
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   5
         Left            =   3600
         TabIndex        =   101
         Top             =   1440
         Width           =   2490
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   4
         Left            =   3600
         TabIndex        =   100
         Top             =   1200
         Width           =   2490
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   3
         Left            =   3600
         TabIndex        =   99
         Top             =   960
         Width           =   2490
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   2
         Left            =   3600
         TabIndex        =   98
         Top             =   720
         Width           =   2490
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   1
         Left            =   3600
         TabIndex        =   97
         Top             =   480
         Width           =   2490
      End
      Begin VB.Label lblPsyDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore 70 HP."
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
         Height          =   210
         Index           =   0
         Left            =   3600
         TabIndex        =   96
         Top             =   240
         Width           =   2490
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   9
         Left            =   3000
         Picture         =   "frmBattle2.frx":3248A
         Top             =   2445
         Width           =   90
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   8
         Left            =   3000
         Picture         =   "frmBattle2.frx":327CC
         Top             =   2205
         Width           =   90
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   7
         Left            =   3000
         Picture         =   "frmBattle2.frx":32B0E
         Top             =   1965
         Width           =   90
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   6
         Left            =   3000
         Picture         =   "frmBattle2.frx":32E50
         Top             =   1725
         Width           =   90
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   5
         Left            =   3000
         Picture         =   "frmBattle2.frx":33192
         Top             =   1485
         Width           =   90
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   4
         Left            =   3000
         Picture         =   "frmBattle2.frx":334D4
         Top             =   1245
         Width           =   90
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   3
         Left            =   3000
         Picture         =   "frmBattle2.frx":33816
         Top             =   1005
         Width           =   90
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   2
         Left            =   3000
         Picture         =   "frmBattle2.frx":33B58
         Top             =   765
         Width           =   90
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   1
         Left            =   3000
         Picture         =   "frmBattle2.frx":33E9A
         Top             =   525
         Width           =   90
      End
      Begin VB.Image imgPsyType 
         Height          =   90
         Index           =   0
         Left            =   3000
         Picture         =   "frmBattle2.frx":341DC
         Top             =   285
         Width           =   90
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   9
         Left            =   2640
         TabIndex        =   94
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   8
         Left            =   2640
         TabIndex        =   93
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   7
         Left            =   2640
         TabIndex        =   92
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   6
         Left            =   2640
         TabIndex        =   91
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   5
         Left            =   2640
         TabIndex        =   90
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   4
         Left            =   2640
         TabIndex        =   89
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   3
         Left            =   2640
         TabIndex        =   88
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   2
         Left            =   2640
         TabIndex        =   87
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   1
         Left            =   2640
         TabIndex        =   86
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblPsyPP 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Height          =   240
         Index           =   0
         Left            =   2640
         TabIndex        =   85
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   9
         Left            =   2280
         TabIndex        =   84
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   8
         Left            =   2280
         TabIndex        =   83
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   7
         Left            =   2280
         TabIndex        =   82
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   6
         Left            =   2280
         TabIndex        =   81
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   5
         Left            =   2280
         TabIndex        =   80
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   4
         Left            =   2280
         TabIndex        =   79
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   3
         Left            =   2280
         TabIndex        =   78
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   2
         Left            =   2280
         TabIndex        =   77
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   1
         Left            =   2280
         TabIndex        =   76
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblPPlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
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
         Height          =   240
         Index           =   0
         Left            =   2280
         TabIndex        =   75
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   8
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":3451E
         MousePointer    =   99  'Custom
         TabIndex        =   73
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   7
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":34DE8
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   6
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":356B2
         MousePointer    =   99  'Custom
         TabIndex        =   71
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   5
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":35F7C
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   4
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":36846
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   3
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":37110
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   2
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":379DA
         MousePointer    =   99  'Custom
         TabIndex        =   67
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   1
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":382A4
         MousePointer    =   99  'Custom
         TabIndex        =   66
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   0
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":38B6E
         MousePointer    =   99  'Custom
         TabIndex        =   65
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   9
         Left            =   120
         Picture         =   "frmBattle2.frx":39438
         Top             =   2400
         Width           =   195
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   8
         Left            =   120
         Picture         =   "frmBattle2.frx":39587
         Top             =   2160
         Width           =   195
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   7
         Left            =   120
         Picture         =   "frmBattle2.frx":396D6
         Top             =   1920
         Width           =   195
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   6
         Left            =   120
         Picture         =   "frmBattle2.frx":39825
         Top             =   1680
         Width           =   195
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   5
         Left            =   120
         Picture         =   "frmBattle2.frx":39974
         Top             =   1440
         Width           =   195
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   4
         Left            =   120
         Picture         =   "frmBattle2.frx":39AC3
         Top             =   1200
         Width           =   195
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   3
         Left            =   120
         Picture         =   "frmBattle2.frx":39C12
         Top             =   960
         Width           =   195
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   2
         Left            =   120
         Picture         =   "frmBattle2.frx":39D61
         Top             =   720
         Width           =   195
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   1
         Left            =   120
         Picture         =   "frmBattle2.frx":39EB0
         Top             =   480
         Width           =   195
      End
      Begin VB.Image imgPsyIcon 
         Height          =   180
         Index           =   0
         Left            =   120
         Picture         =   "frmBattle2.frx":39FFF
         Top             =   240
         Width           =   195
      End
      Begin VB.Label lblPsynergy 
         BackStyle       =   0  'Transparent
         Caption         =   "Cure"
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
         Height          =   240
         Index           =   9
         Left            =   360
         MouseIcon       =   "frmBattle2.frx":3A14E
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   2400
         Width           =   1695
      End
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Commands"
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
      Height          =   585
      Left            =   0
      TabIndex        =   193
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgBattleIcon 
      Height          =   720
      Index           =   7
      Left            =   6000
      MouseIcon       =   "frmBattle2.frx":3AA18
      MousePointer    =   99  'Custom
      Picture         =   "frmBattle2.frx":3B2E2
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image imgChooseEnemy 
      Height          =   360
      Left            =   1680
      Picture         =   "frmBattle2.frx":3B508
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
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
      Height          =   480
      Left            =   6840
      TabIndex        =   41
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Shape shpDesc 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00886000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   4
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Image imgFace 
      Height          =   840
      Left            =   120
      Picture         =   "frmBattle2.frx":3B874
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   840
   End
   Begin VB.Label lblChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Felix"
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
      Height          =   240
      Index           =   3
      Left            =   6480
      TabIndex        =   19
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Felix"
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
      Height          =   240
      Index           =   2
      Left            =   4440
      TabIndex        =   18
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Felix"
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
      Height          =   240
      Index           =   1
      Left            =   2400
      TabIndex        =   17
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Felix"
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
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblPP 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   240
      Index           =   3
      Left            =   6840
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblPP 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   240
      Index           =   2
      Left            =   4800
      TabIndex        =   14
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblPP 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   240
      Index           =   1
      Left            =   2760
      TabIndex        =   13
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblPP 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   240
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   240
      Index           =   3
      Left            =   6840
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   240
      Index           =   2
      Left            =   4800
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   240
      Index           =   1
      Left            =   2760
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Height          =   240
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PP"
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
      Height          =   240
      Index           =   7
      Left            =   6360
      TabIndex        =   7
      Top             =   720
      Width           =   270
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PP"
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
      Height          =   240
      Index           =   6
      Left            =   4320
      TabIndex        =   6
      Top             =   720
      Width           =   270
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PP"
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
      Height          =   240
      Index           =   5
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   270
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PP"
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
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   270
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
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
      Height          =   240
      Index           =   3
      Left            =   6360
      TabIndex        =   3
      Top             =   240
      Width           =   270
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
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
      Height          =   240
      Index           =   2
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   270
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
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
      Height          =   240
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   270
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
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
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   270
   End
   Begin VB.Shape shpPP 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape shpPPB 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape shpPP 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape shpPPB 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape shpPP 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape shpPPB 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape shpPP 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape shpPPB 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape shpHP 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape shpHPB 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape shpHP 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape shpHPB 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape shpHP 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape shpHPB 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape shpHP 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape shpHPB 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Enemy 
      Height          =   960
      Index           =   7
      Left            =   6360
      MouseIcon       =   "frmBattle2.frx":3BB52
      Picture         =   "frmBattle2.frx":3C41C
      Top             =   3480
      Width           =   960
   End
   Begin VB.Image Enemy 
      Height          =   960
      Index           =   6
      Left            =   5280
      MouseIcon       =   "frmBattle2.frx":3D15D
      Picture         =   "frmBattle2.frx":3DA27
      Top             =   3480
      Width           =   960
   End
   Begin VB.Image Enemy 
      Height          =   960
      Index           =   5
      Left            =   4320
      MouseIcon       =   "frmBattle2.frx":3E8CF
      Picture         =   "frmBattle2.frx":3F199
      Top             =   3480
      Width           =   960
   End
   Begin VB.Image Enemy 
      Height          =   960
      Index           =   4
      Left            =   3240
      MouseIcon       =   "frmBattle2.frx":3F8BC
      Picture         =   "frmBattle2.frx":40186
      Top             =   3480
      Width           =   960
   End
   Begin VB.Image Enemy 
      Height          =   960
      Index           =   3
      Left            =   4320
      MouseIcon       =   "frmBattle2.frx":410E5
      Picture         =   "frmBattle2.frx":419AF
      Top             =   2880
      Width           =   960
   End
   Begin VB.Image Enemy 
      Height          =   960
      Index           =   2
      Left            =   3240
      MouseIcon       =   "frmBattle2.frx":420A7
      Picture         =   "frmBattle2.frx":42971
      Top             =   2880
      Width           =   960
   End
   Begin VB.Image Enemy 
      Height          =   960
      Index           =   1
      Left            =   2160
      MouseIcon       =   "frmBattle2.frx":4311E
      Picture         =   "frmBattle2.frx":439E8
      Top             =   2880
      Width           =   960
   End
   Begin VB.Image Enemy 
      Height          =   960
      Index           =   0
      Left            =   1200
      MouseIcon       =   "frmBattle2.frx":44C0C
      Picture         =   "frmBattle2.frx":454D6
      Top             =   2880
      Width           =   960
   End
   Begin VB.Image imgBG 
      Height          =   3000
      Left            =   1200
      Picture         =   "frmBattle2.frx":46884
      Top             =   1440
      Width           =   6600
   End
   Begin VB.Shape shpStats 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00886000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Index           =   3
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2055
   End
   Begin VB.Shape shpStats 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00886000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Index           =   2
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2055
   End
   Begin VB.Shape shpStats 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00886000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Index           =   1
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2055
   End
   Begin VB.Shape shpStats 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00886000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Index           =   0
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image imgBattleIcon 
      Height          =   720
      Index           =   6
      Left            =   5280
      MouseIcon       =   "frmBattle2.frx":51095
      MousePointer    =   99  'Custom
      Picture         =   "frmBattle2.frx":5195F
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image imgBattleIcon 
      Height          =   720
      Index           =   5
      Left            =   4560
      MouseIcon       =   "frmBattle2.frx":51B3E
      MousePointer    =   99  'Custom
      Picture         =   "frmBattle2.frx":52408
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image imgBattleIcon 
      Height          =   720
      Index           =   4
      Left            =   3840
      MouseIcon       =   "frmBattle2.frx":52654
      MousePointer    =   99  'Custom
      Picture         =   "frmBattle2.frx":52F1E
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image imgBattleIcon 
      Height          =   720
      Index           =   3
      Left            =   3120
      MouseIcon       =   "frmBattle2.frx":530EB
      MousePointer    =   99  'Custom
      Picture         =   "frmBattle2.frx":539B5
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image imgBattleIcon 
      Height          =   720
      Index           =   2
      Left            =   2400
      MouseIcon       =   "frmBattle2.frx":53BB3
      MousePointer    =   99  'Custom
      Picture         =   "frmBattle2.frx":5447D
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image imgBattleIcon 
      Height          =   720
      Index           =   1
      Left            =   1680
      MouseIcon       =   "frmBattle2.frx":546B5
      MousePointer    =   99  'Custom
      Picture         =   "frmBattle2.frx":54F7F
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image imgBattleIcon 
      Height          =   720
      Index           =   0
      Left            =   960
      MouseIcon       =   "frmBattle2.frx":551E8
      MousePointer    =   99  'Custom
      Picture         =   "frmBattle2.frx":55AB2
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   720
   End
End
Attribute VB_Name = "frmBattle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCommand As Long
Dim bChooseAlly As Boolean
Dim bChooseEnemy As Boolean
Dim intCurStatusChar As Long
Dim intCurSelected As Long

Private Sub Enemy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If bChooseEnemy = True Then
    curEnemy = Index
    bChooseEnemy = False
    imgChooseEnemy.Visible = False
    For i = 0 To 3
        Enemy(i).MousePointer = 0
    Next 'i
    For i = 0 To imgBattleIcon.UBound
        imgBattleIcon(i).Visible = True
    Next 'i
    If intCommand = 0 Then
        Call subAttack
    ElseIf intCommand = 1 Then
        Call subPsynergy
    ElseIf intCommand = 2 Then
        Call subDjinn
    ElseIf intCommand = 3 Then
        Call subSummon
    ElseIf intCommand = 4 Then
        Call subItem
    End If
End If
End Sub

Private Sub Enemy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If bChooseEnemy = True And Index < 4 Then
    imgChooseEnemy.Left = Enemy(Index).Left + (Enemy(Index).Width / 2 - 10)
End If
If bChooseAlly = True And Index >= 4 Then
    imgChooseEnemy.Left = Enemy(Index).Left + (Enemy(Index).Width / 2 - 10)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
'Comment out in actual version
curChar = 1
WOTAChar(1).Level = 10
WOTAChar(1).Type = "Earth"
WOTAChar(1).Name = "Felix"
WOTAChar(1).Num = 5
For i = 1 To 5
    WOTAChar(1).DjinnNum(i) = i
    WOTAChar(1).DjinnEnabled(i) = True
Next 'i
Call GetDjinn(curChar)
Call GetClass(curChar)

WOTAChar(2).Level = 10
WOTAChar(2).Type = "Fire"
WOTAChar(2).Name = "Garet"
WOTAChar(2).Num = 2
For i = 1 To 5
    WOTAChar(2).DjinnNum(i) = i
    WOTAChar(2).DjinnEnabled(i) = True
Next 'i
Call GetDjinn(curChar)
Call GetClass(curChar)

WOTAChar(3).Level = 10
WOTAChar(3).Type = "Earth"
WOTAChar(3).Name = "Isaac"
WOTAChar(3).Num = 1
For i = 1 To 5
    WOTAChar(3).DjinnNum(i) = i
    WOTAChar(3).DjinnEnabled(i) = True
Next 'i
Call GetDjinn(curChar)
Call GetClass(curChar)

WOTAChar(4).Level = 10
WOTAChar(4).Type = "Water"
WOTAChar(4).Name = "Mia"
WOTAChar(4).Num = 4
For i = 1 To 5
    WOTAChar(4).DjinnNum(i) = i
    WOTAChar(4).DjinnEnabled(i) = True
Next 'i
Call GetDjinn(curChar)
Call GetClass(curChar)

Call LoadNewBattle
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub imgBattleIcon_Click(Index As Integer)
Select Case Index
Case 0
    intCommand = 0
    Call ChooseEnemy
Case 1
    Call subPsynergy
Case 2
    Call subDisplayDjinn
Case 3
    Call subSummon
Case 4
    Call subItem
Case 5
    Call subDefend
Case 6
    Call subStatus
Case 7
    Call subBack
End Select
End Sub
Sub subAttack()
On Error Resume Next

Dim RelativeDamage As Long 'Ammount of physical damage to be done
Dim RelativePower As Long
Dim intRand As Long

RelativeDamage = WOTAChar(curChar).AP - WOTAChar(curEnemy).Defense
If RelativeDamage < 0 Then RelativeDamage = 0
RelativeDamage = RelativeDamage / 2
Select Case WOTAChar(curEnemy).Type
Case "Earth"
    RelativePower = WOTAChar(curChar).EarthPower - WOTAChar(curEnemy).EarthResist
Case "Fire"
    RelativePower = WOTAChar(curChar).FirePower - WOTAChar(curEnemy).FireResist
Case "Wind"
    RelativePower = WOTAChar(curChar).WindPower - WOTAChar(curEnemy).WindResist
Case Else
    RelativePower = WOTAChar(curChar).WaterPower - WOTAChar(curEnemy).WaterResist
End Select
If RelativePower < -200 Then RelativePower = -200
If RelativePower > 200 Then RelativePower = 200

Randomize
intRand = Int(Rnd * 100) + 1
WOTAChar(curChar).Command = "ATTACK"
WOTAChar(curChar).Target = curEnemy

If intRand <= WOTAChar(curChar).ItemNum Then
    WOTAChar(curChar).Command = "ATTACKS"
    RelativeDamage = RelativeDamage + WOTAChar(curChar).ItemNum
    RelativeDamage = RelativeDamage * (1 + (RelativePower / 400))
ElseIf intRand <= WOTAChar(curChar).Luck Then
    WOTAChar(curChar).Command = "ATTACKC"
    RelativeDamage = RelativeDamage * 1.25
    RelativeDamage = RelativeDamage + (WOTAChar(curEnemy).Level / 5) + 6
End If

RelativeDamage = RelativeDamage + Int(Rnd * 4)
WOTAChar(curChar).Damage = RelativeDamage

Call NextCharacter

End Sub
Sub subPsynergy()
On Error Resume Next
Dim curPsy As Long
curPsy = 0
For i = 0 To 9
    lblPsynergy(i).Visible = False
    lblPsyPP(i).Visible = False
    lblPsyDesc(i).Visible = False
    imgPsyIcon(i).Visible = False
    imgPsyRange(i).Visible = False
    imgPsyType(i).Visible = False
    lblPPlbl(i).Visible = False
Next 'i
Dim strSplit
Dim strImage As String
With WOTAChar(curChar)
    For i = 1 To 100
        For q = 1 To 20
            If nPsynergy(i).ClassName(q) <> "" Then
                For w = 1 To 10
                    If nClass(.ClassNum).ClassInherit(w) <> "" Then
                        If nClass(.ClassNum).ClassInherit(w) = nPsynergy(i).ClassName(q) Then
                            If nPsynergy(i).ClassLVL(q) <= .Level Then
                                lblPsynergy(curPsy).Caption = nPsynergy(i).Name
                                lblPP(curPsy).Caption = nPsynergy(i).PP
                                lblPsyDesc(curPsy).Caption = nPsynergy(i).Description
                                lblPsynergy(curPsy).Visible = True
                                lblPsyPP(curPsy).Visible = True
                                lblPsyDesc(curPsy).Visible = True
                                imgPsyIcon(curPsy).Visible = True
                                imgPsyRange(curPsy).Visible = True
                                imgPsyType(curPsy).Visible = True
                                lblPPlbl(curPsy).Visible = True
                                strImage = ""
                                strSplit = Split(nPsynergy(i).Name, " ", -1, vbTextCompare)
                                For e = 0 To UBound(strSplit)
                                    If e = 0 Then
                                        strImage = strImage & strSplit(e)
                                    Else
                                        strImage = strImage & "_" & strSplit(e)
                                    End If
                                Next 'e
                                strImage = App.Path & "\BattleImages\PsyIcons\" & strImage & ".gif"
                                imgPsyIcon(curPsy).Picture = LoadPicture(strImage)
                                imgPsyRange(curPsy).Picture = LoadPicture(App.Path & "\BattleImages\range-" & CStr(nPsynergy(i).Range) & ".gif")
                                curPsy = curPsy + 1
                            End If
                        End If
                    End If
                Next 'w
                If nPsynergy(i).ClassName(q) = .ClassName Then
                    If nPsynergy(i).ClassLVL(q) <= .Level Then
                        lblPsynergy(curPsy).Caption = nPsynergy(i).Name
                        lblPP(curPsy).Caption = nPsynergy(i).PP
                        lblPsyDesc(curPsy).Caption = nPsynergy(i).Description
                        lblPsynergy(curPsy).Visible = True
                        lblPsyPP(curPsy).Visible = True
                        lblPsyDesc(curPsy).Visible = True
                        imgPsyIcon(curPsy).Visible = True
                        imgPsyRange(curPsy).Visible = True
                        imgPsyType(curPsy).Visible = True
                        lblPPlbl(curPsy).Visible = True
                        strImage = ""
                        strSplit = Split(nPsynergy(i).Name, " ", -1, vbTextCompare)
                        For e = 0 To UBound(strSplit)
                            If e = 0 Then
                                strImage = strImage & strSplit(e)
                            Else
                                strImage = strImage & "_" & strSplit(e)
                            End If
                        Next 'e
                        strImage = App.Path & "\BattleImages\PsyIcons\" & strImage & ".gif"
                        imgPsyIcon(curPsy).Picture = LoadPicture(strImage)
                        imgPsyRange(curPsy).Picture = LoadPicture(App.Path & "\BattleImages\range-" & CStr(nPsynergy(i).Range) & ".gif")
                        curPsy = curPsy + 1
                    End If
                End If
            End If
        Next 'q
    Next 'i
End With

framPsynergy.Visible = True

End Sub
Sub subDjinn()
On Error Resume Next
Dim RelativeDamage As Long 'Ammount of physical damage to be done
Dim RelativePower As Long
Dim curDjinn As Long



RelativeDamage = WOTAChar(curChar).AP - WOTAChar(curEnemy).Defense
If RelativeDamage < 0 Then RelativeDamage = 0
RelativeDamage = RelativeDamage / 2

Select Case nDjinn(curDjinn).Element
Case "Earth"
    RelativePower = WOTAChar(curChar).EarthPower - WOTAChar(curEnemy).EarthResist
Case "Fire"
    RelativePower = WOTAChar(curChar).FirePower - WOTAChar(curEnemy).FireResist
Case "Wind"
    RelativePower = WOTAChar(curChar).WindPower - WOTAChar(curEnemy).WindResist
Case Else
    RelativePower = WOTAChar(curChar).WaterPower - WOTAChar(curEnemy).WaterResist
End Select
If RelativePower < -200 Then RelativePower = -200
If RelativePower > 200 Then RelativePower = 200

RelativeDamage = RelativeDamage * (1 + (RelativePower / 400))

Select Case nDjinn(curDjinn).Type
Case "ADAMAGE" 'Djinn does attack + add mod
    RelativeDamage = RelativeDamage + nDjinn(curDjinn).Damage
Case "MDAMAGE" 'Djinn does attack * mult mod
    RelativeDamage = RelativeDamage * (nDjinn(curDjinn).Damage / 100)
End Select

End Sub
Sub subSummon()
On Error Resume Next
Dim RelativePower As Long

Dim RelativeDamage As Long
Dim curSummon As Long
Dim DjinnUsed As Long

Select Case nSummon(curSummon).Element
Case "Earth"
    RelativePower = WOTAChar(curChar).EarthPower - WOTAChar(curEnemy).EarthResist
Case "Fire"
    RelativePower = WOTAChar(curChar).FirePower - WOTAChar(curEnemy).FireResist
Case "Wind"
    RelativePower = WOTAChar(curChar).WindPower - WOTAChar(curEnemy).WindResist
Case "Water"
    RelativePower = WOTAChar(curChar).WaterPower - WOTAChar(curEnemy).WaterResist
End Select

With nSummon(curSummon)
    DjinnUsed = .EarthDjinn + .FireDjinn + .WaterDjinn + .WindDjinn
End With

RelativeDamage = nSummon(curSummon).BaseDamage
RelativeDamage = RelativeDamage + (WOTAChar(curEnemy).MaxHP * 3 * DjinnUsed / 100)
RelativeDamage = RelativeDamage * (1 + (RelativePower / 200))

End Sub
Sub subStatus()
On Error Resume Next
intCurStatusChar = curChar
Call DisplayStatus
End Sub

Private Sub imgBattleIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
    lblDesc.Caption = "Attack"
Case 1
    lblDesc.Caption = "Psynergy"
Case 2
    lblDesc.Caption = "Djinn"
Case 3
    lblDesc.Caption = "Summon"
Case 4
    lblDesc.Caption = "Item"
Case 5
    lblDesc.Caption = "Defend"
Case 6
    lblDesc.Caption = "Status"
Case 7
    lblDesc.Caption = "Back"
End Select

End Sub

Private Sub lblDjinnDesc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
For i = 1 To 72
    'nothing right now
Next 'i
End Sub

Private Sub lblPsynergy_Click(Index As Integer)
framPsynergy.Visible = False
intCommand = 1
Call ChooseEnemy
End Sub

Private Sub lblSSwitchL_Click()
On Error Resume Next
intCurStatusChar = intCurStatusChar - 1
If intCurStatusChar < 1 Then intCurStatusChar = 4
Call DisplayStatus
End Sub

Private Sub lblSSwitchR_Click()
On Error Resume Next
intCurStatusChar = intCurStatusChar + 1
If intCurStatusChar > 4 Then intCurStatusChar = 1
Call DisplayStatus

End Sub
Sub DisplayStatus()
On Error Resume Next
Me.lblSChar.Caption = WOTAChar(intCurStatusChar).Name
Me.lblSClass.Caption = WOTAChar(intCurStatusChar).ClassName
'Me.lblSExp.Caption = WOTAChar(intCurStatusChar).Exp
Me.lblSLvl.Caption = WOTAChar(intCurStatusChar).Level
Me.lblSStatus.Caption = WOTAChar(intCurStatusChar).Status
Me.lblSAgility.Caption = WOTAChar(intCurStatusChar).Agility
Me.lblSAttack.Caption = WOTAChar(intCurStatusChar).AP
Me.lblSDefense.Caption = WOTAChar(intCurStatusChar).Defense
Me.lblSHP.Caption = WOTAChar(intCurStatusChar).HP
Me.lblSPP.Caption = WOTAChar(intCurStatusChar).PP
Me.lblSLuck.Caption = WOTAChar(intCurStatusChar).Luck
Me.lblSEDjinn.Caption = WOTAChar(intCurStatusChar).EarthDjinn
Me.lblSELevel.Caption = WOTAChar(intCurStatusChar).EarthLevel
Me.lblSEPower.Caption = WOTAChar(intCurStatusChar).EarthPower
Me.lblSEResist.Caption = WOTAChar(intCurStatusChar).EarthResist
Me.lblSFDjinn.Caption = WOTAChar(intCurStatusChar).FireDjinn
Me.lblSFLevel.Caption = WOTAChar(intCurStatusChar).FireLevel
Me.lblSFPower.Caption = WOTAChar(intCurStatusChar).FirePower
Me.lblSFResist.Caption = WOTAChar(intCurStatusChar).FireResist
Me.lblSNDjinn.Caption = WOTAChar(intCurStatusChar).WindDjinn
Me.lblSNLevel.Caption = WOTAChar(intCurStatusChar).WindLevel
Me.lblSNPower.Caption = WOTAChar(intCurStatusChar).WindPower
Me.lblSNResist.Caption = WOTAChar(intCurStatusChar).WindResist
Me.lblSWDjinn.Caption = WOTAChar(intCurStatusChar).WaterDjinn
Me.lblSWLevel.Caption = WOTAChar(intCurStatusChar).WaterLevel
Me.lblSWPower.Caption = WOTAChar(intCurStatusChar).WaterPower
Me.lblSWResist.Caption = WOTAChar(intCurStatusChar).WaterResist
Me.imgSPic.Picture = LoadPicture(App.Path & "\BattleImages\" & WOTAChar(intCurStatusChar).Name & "B.gif")
Me.imgSFace.Picture = LoadPicture(App.Path & "\BattleImages\CharIcons\" & WOTAChar(intCurStatusChar).Name & ".gif")
Me.lblSWeapon.Caption = WOTAChar(intCurStatusChar).WeaponName
Me.lblSArmor(0).Caption = WOTAChar(intCurStatusChar).ArmorChestName
Me.lblSArmor(1).Caption = WOTAChar(intCurStatusChar).ArmorArmName
Me.lblSArmor(2).Caption = WOTAChar(intCurStatusChar).ArmorMiscName

framStatus.Visible = True
End Sub
Sub ChooseEnemy()
On Error Resume Next
bChooseEnemy = True
imgChooseEnemy.Visible = True
lblDesc.Caption = "Choose Enemy"
For i = 0 To 3
    Enemy(i).MousePointer = 99
Next 'i
For i = 0 To imgBattleIcon.UBound
    imgBattleIcon(i).Visible = False
Next 'i

End Sub
Sub subItem()
'nothing
End Sub

Private Sub txtChat_Change()
On Error Resume Next
Call AutoScrollTxt(txtChat)
End Sub

Sub DoCommands()
On Error Resume Next
curChar = FindFirstAttack
If curChar = 0 Then Exit Sub 'No one left to attack

If curChar < 4 Then 'If the character is on your team
    curEnemy = WOTAChar(curChar).Target
    Select Case WOTAChar(curChar).Command
        Case 0
            Call subAttack
        Case 1
            Call subPsynergy
        Case 2
            Call subDjinn
        Case 3
            Call subSummon
        Case 4
            Call subItem
        Case 5
            Call subDefend
    End Select
    WOTAChar(curChar).DidMove = True
End If
    
    
End Sub
Function FindFirstAttack() As Long
On Error Resume Next
Dim curHighestAgility As Long
Dim curHighestChar As Long
For i = 1 To 8
    If WOTAChar(i).DidMove = False And WOTAChar(i).Enabled = True And WOTAChar(i).HP > 0 Then
        If WOTAChar(i).Agility > curHighestAgility Then
            curHighestAgility = WOTAChar(i).Agility
            curHighestChar = i
        End If
    End If
Next 'i
FindFirstAttack = curHighestChar
End Function
Sub subDefend()
'Defend
On Error Resume Next
WOTAChar(curChar).Command = "DEFEND"
Call NextCharacter
End Sub

Private Sub txtChatMsg_Change()
If txtChatMsg.Text = "loadbattle" Then
    Call LoadNewBattle
    txtChatMsg.Text = ""
End If
End Sub

Private Sub txtChatMsg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    On Error Resume Next
    Call SendBattleData("BATTLECHAT" & strMyUserName & ": " & txtChatMsg.Text & vbCrLf)
    txtChat.Text = txtChat.Text & vbNewLine & strMyUserName & ": " & txtChatMsg.Text
    txtChatMsg.Text = ""
End If
If KeyCode = vbKeyEscape And Shift = 1 Then
    Unload Me
End If
End Sub
Sub GetClass(intChar As Long)
On Error Resume Next
With WOTAChar(intChar)
    For i = 1 To 175
        If (.Type = "Earth" And nClass(i).Earth = True) Or (.Type = "Fire" And nClass(i).Fire = True) Or (.Type = "Wind" And nClass(i).Wind = True) Or (.Type = "Water" And nClass(i).Water = True) Then
            If .EarthDjinn >= nClass(i).EarthMin And .EarthDjinn <= nClass(i).EarthMax And .FireDjinn >= nClass(i).FireMin And .FireDjinn <= nClass(i).FireMax And .WindDjinn >= nClass(i).WindMin And .WindDjinn <= nClass(i).WindMax And .WaterDjinn >= nClass(i).WaterMin And .WaterDjinn <= nClass(i).WaterMax Then
                .ClassName = nClass(i).Name
                .ClassNum = i
                .EarthLevel = nClass(i).EarthLVL
                .FireLevel = nClass(i).FireLVL
                .WindLevel = nClass(i).WindLVL
                .WaterLevel = nClass(i).WaterLVL
                Exit For
            End If
        End If
    Next 'i
End With

End Sub
Sub GetDjinn(intChar As Long)
WOTAChar(intChar).EarthDjinn = 0
WOTAChar(intChar).FireDjinn = 0
WOTAChar(intChar).WindDjinn = 0
WOTAChar(intChar).WaterDjinn = 0
For i = 1 To 9
    If WOTAChar(intChar).DjinnNum(i) <> 0 Then
        With nDjinn(WOTAChar(intChar).DjinnNum(i))
            If WOTAChar(intChar).DjinnState(i) = 0 And WOTAChar(intChar).DjinnEnabled(i) = True Then
                If .Element = "Earth" Then
                    WOTAChar(intChar).EarthDjinn = WOTAChar(intChar).EarthDjinn + 1
                ElseIf .Element = "Fire" Then
                    WOTAChar(intChar).FireDjinn = WOTAChar(intChar).FireDjinn + 1
                ElseIf .Element = "Wind" Then
                    WOTAChar(intChar).WindDjinn = WOTAChar(intChar).WindDjinn + 1
                ElseIf .Element = "Water" Then
                    WOTAChar(intChar).WaterDjinn = WOTAChar(intChar).WaterDjinn + 1
                End If
            End If
        End With
    End If
Next 'i

End Sub
Sub subBack()
On Error Resume Next
'Goes back to the previous character's turn
If framStatus.Visible = True Then
    framStatus.Visible = False
ElseIf framPsynergy.Visible = True Then
    framPsynergy.Visible = False
ElseIf framItems.Visible = True Then
    framItems.Visible = False
ElseIf framDjinn.Visible = True Then
    framDjinn.Visible = False
ElseIf framSummon.Visible = True Then
    framSummon.Visible = False
ElseIf curChar > 1 Then
    curChar = curChar - 1
    imgFace.Picture = LoadPicture(App.Path & "\BattleImages\CharIcons\" & WOTAChar(curChar).Name & ".gif")
End If
End Sub
Sub subPsyDamage(curPsynergy As Long)
Dim BaseDamage As Long
Dim RelativePower As Long
Dim intRand As Long
Randomize
intRand = Int(Rnd * 100) + 1
Select Case nPsynergy(curPsynergy).Type
    Case "Regular Magic"
        Select Case nPsynergy(curPsynergy).Element
        Case "Earth"
            RelativePower = WOTAChar(curChar).EarthPower - WOTAChar(curEnemy).EarthResist
        Case "Fire"
            RelativePower = WOTAChar(curChar).FirePower - WOTAChar(curEnemy).FireResist
        Case "Wind"
            RelativePower = WOTAChar(curChar).WindPower - WOTAChar(curEnemy).WindResist
        Case Else
            RelativePower = WOTAChar(curChar).WaterPower - WOTAChar(curEnemy).WaterResist
        End Select
        If RelativePower < -200 Then RelativePower = -200
        If RelativePower > 200 Then RelativePower = 200
        BaseDamage = nPsynergy(curPsynergy).Damage
        BaseDamage = BaseDamage * (1 + (RelativePower / 200))
        WOTAChar(curChar).Target = curEnemy
        WOTAChar(curChar).Command = "PSY"
        WOTAChar(curChar).Damage = BaseDamage
    Case "Curative"
        Select Case nPsynergy(curPsynergy).Element
        Case "Earth"
            RelativePower = WOTAChar(curChar).EarthPower
        Case "Fire"
            RelativePower = WOTAChar(curChar).FirePower
        Case "Wind"
            RelativePower = WOTAChar(curChar).WindPower
        Case Else
            RelativePower = WOTAChar(curChar).WaterPower
        End Select
        BaseDamage = nPsynergy(curPsynergy).Damage
        BaseDamage = BaseDamage * (RelativePower / 100)
        WOTAChar(curChar).Target = curEnemy
        WOTAChar(curChar).Command = "PSYH"
        WOTAChar(curChar).Damage = BaseDamage
    Case "Revive"
        WOTAChar(curChar).Target = curEnemy
        If WOTAChar(curChar).Luck >= intRand Then
            WOTAChar(curChar).Command = "PSYREVIVES"
        Else
            WOTAChar(curChar).Command = "PSYREVIVEF"
        End If
    Case "Cure Poison"
        WOTAChar(curChar).Target = curEnemy
        WOTAChar(curChar).Command = "PSYPOISON"
    Case "Restore"
        WOTAChar(curChar).Target = curEnemy
        WOTAChar(curChar).Command = "PSYRESTORE"
    Case "Attack Plus Add Mod"
        'nothing
    Case "Attack Plus Mult Mod"
        'nothing
    
End Select

End Sub
Sub NextCharacter()
curChar = curChar + 1
If curChar > 4 Then
    Call FinishTurn
Else
    imgFace.Picture = LoadPicture(App.Path & "\BattleImages\CharIcons\" & WOTAChar(curChar).Name & ".gif")
End If
    
End Sub
Sub FinishTurn()
On Error Resume Next
For i = 0 To imgBattleIcon.UBound
    imgBattleIcon(i).Visible = False
Next 'i
imgFace.Visible = False
shpDesc(4).Visible = False
lblDesc.Visible = False
bPlayer1Ready = True
Call SendBattleData("READY")
If bPlayer2Ready = True Then
    Call DoCommands
Else
    lblTime.Caption = "Waiting for opponent"
End If

For i = 1 To 4
    Call SendBattleData("CHARCOM" & CStr(i) & WOTAChar(i).Command & vbCrLf)
    Call SendBattleData("CHARTARGET" & CStr(i) & WOTAChar(i).Target & vbCrLf)
    Call SendBattleData("CHARDAMAGE" & CStr(i) & WOTAChar(i).Damage & vbCrLf)
Next 'i
Call SendBattleData("FINISHTURN" & vbCrLf)

txtChat.Text = txtChat.Text & vbNewLine & "You have selected your commands."
End Sub
Sub subDisplayDjinn()
On Error Resume Next
framDjinn.Visible = True
Me.lblDOClass.Caption = WOTAChar(curChar).ClassName
Me.lblDOAgility.Caption = WOTAChar(curChar).Agility
Me.lblDOAttack.Caption = WOTAChar(curChar).AP
Me.lblDODefense.Caption = WOTAChar(curChar).Defense
Me.lblDOHP.Caption = WOTAChar(curChar).HP
Me.lblDOLuck.Caption = WOTAChar(curChar).Luck
Me.lblDOPP.Caption = WOTAChar(curChar).PP
Dim curDjinn As Long
curDjinn = 0
For i = 1 To 9
    If WOTAChar(curChar).DjinnEnabled(i) = True Then
        With nDjinn(WOTAChar(curChar).DjinnNum(i))
            Me.lblDjinn(curDjinn).Caption = .Name
            Me.lblDjinnDesc(curDjinn).Caption = .Description
            curDjinn = curDjinn + 1
        End With
    End If
Next 'i
End Sub
Sub Reset()
'Loads the first character
On Error Resume Next
For i = 0 To imgBattleIcon.UBound
    imgBattleIcon(i).Visible = True
Next 'i
framStatus.Visible = False
framPsynergy.Visible = False
framItems.Visible = False
framDjinn.Visible = False
framSummon.Visible = False
lblTime.Caption = "Select your commands."
imgFace.Picture = LoadPicture(App.Path & "\BattleImages\CharIcons\" & WOTAChar(curChar).Name & ".gif")
curChar = 1
End Sub
