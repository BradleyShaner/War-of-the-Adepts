VERSION 5.00
Begin VB.Form frmBattle 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   ControlBox      =   0   'False
   Icon            =   "frmBattle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   39
      Top             =   4620
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.Timer timeLoadBattle 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   6960
      Top             =   1200
   End
   Begin VB.TextBox txtDisplay 
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
      Height          =   1215
      Left            =   2040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   4560
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Timer timeWait 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4320
      Top             =   1320
   End
   Begin VB.ListBox lstSummon 
      Height          =   1035
      ItemData        =   "frmBattle.frx":08CA
      Left            =   5160
      List            =   "frmBattle.frx":08CC
      TabIndex        =   33
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox lstDjinn 
      Height          =   1035
      ItemData        =   "frmBattle.frx":08CE
      Left            =   5160
      List            =   "frmBattle.frx":08D0
      TabIndex        =   32
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox lstPsynergy 
      Height          =   1035
      ItemData        =   "frmBattle.frx":08D2
      Left            =   5160
      List            =   "frmBattle.frx":08D4
      TabIndex        =   31
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer timeSummon 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   3720
      Top             =   1920
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset (Only Do This If The Game Is Frozen)"
      Height          =   615
      Left            =   7680
      TabIndex        =   24
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Timer timeDjinn 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   3360
      Top             =   1920
   End
   Begin VB.Timer timeSword 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   2640
      Top             =   1920
   End
   Begin VB.Timer timeBoready 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   3240
   End
   Begin VB.Timer timecount 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   600
   End
   Begin VB.Timer timePsynergy 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   3000
      Top             =   1920
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   8280
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtmsg 
      Height          =   285
      Left            =   6240
      MaxLength       =   255
      TabIndex        =   12
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image imgUser 
      Height          =   720
      Index           =   0
      Left            =   0
      Picture         =   "frmBattle.frx":08D6
      Top             =   1920
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblSelectTarget 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Target"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   46
      Top             =   5760
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   6
      Left            =   8400
      Picture         =   "frmBattle.frx":32A0
      Top             =   4920
      Width           =   345
   End
   Begin VB.Label lblBackTurn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   8160
      TabIndex        =   45
      Top             =   5280
      Width           =   810
   End
   Begin VB.Image imgTurn 
      Height          =   240
      Left            =   0
      Picture         =   "frmBattle.frx":36BE
      Top             =   960
      Width           =   240
   End
   Begin VB.Image imgYou 
      Height          =   915
      Index           =   6
      Left            =   3240
      MouseIcon       =   "frmBattle.frx":3A32
      Picture         =   "frmBattle.frx":42FC
      Top             =   3480
      Width           =   525
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   44
      Top             =   1200
      Width           =   855
   End
   Begin VB.Shape shpHP 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblPP 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   43
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "PP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   42
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   41
      Top             =   1320
      Width           =   855
   End
   Begin VB.Shape shpHP 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   855
   End
   Begin VB.Image imgYou 
      Height          =   600
      Index           =   5
      Left            =   2760
      Picture         =   "frmBattle.frx":49DB
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image imgYou 
      Height          =   600
      Index           =   4
      Left            =   120
      Picture         =   "frmBattle.frx":4D84
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image imgUser 
      Height          =   735
      Index           =   23
      Left            =   1080
      Picture         =   "frmBattle.frx":512D
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgUser 
      Height          =   780
      Index           =   22
      Left            =   240
      Picture         =   "frmBattle.frx":56D4
      Top             =   2040
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblStatusClose 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   40
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   5
      Left            =   7320
      Picture         =   "frmBattle.frx":5C6A
      Top             =   4920
      Width           =   360
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S&tatus"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   6960
      TabIndex        =   38
      Top             =   5280
      Width           =   1020
   End
   Begin VB.Image imgUser 
      Height          =   600
      Index           =   99
      Left            =   0
      Picture         =   "frmBattle.frx":6081
      Top             =   2280
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgBack 
      Height          =   375
      Left            =   6960
      Picture         =   "frmBattle.frx":642A
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOK 
      Height          =   375
      Left            =   5160
      Picture         =   "frmBattle.frx":686B
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgUser 
      Height          =   540
      Index           =   21
      Left            =   2160
      Picture         =   "frmBattle.frx":6CA0
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgUser 
      Height          =   735
      Index           =   20
      Left            =   960
      Picture         =   "frmBattle.frx":72B9
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgUser 
      Height          =   735
      Index           =   19
      Left            =   1080
      Picture         =   "frmBattle.frx":79D6
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgUser 
      Height          =   900
      Index           =   18
      Left            =   1560
      Picture         =   "frmBattle.frx":8127
      Top             =   3480
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgUser 
      Height          =   735
      Index           =   17
      Left            =   2400
      Picture         =   "frmBattle.frx":89EC
      Top             =   3000
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgRealPsy 
      Height          =   2400
      Left            =   3000
      Picture         =   "frmBattle.frx":906C
      Top             =   2400
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Image imgUser 
      Height          =   735
      Index           =   16
      Left            =   360
      Picture         =   "frmBattle.frx":A4FB
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgUser 
      Height          =   810
      Index           =   15
      Left            =   240
      Picture         =   "frmBattle.frx":AB8E
      Top             =   3240
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgUser 
      Height          =   900
      Index           =   14
      Left            =   840
      Picture         =   "frmBattle.frx":B179
      Top             =   1680
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   4
      Left            =   6000
      Picture         =   "frmBattle.frx":BBA4
      Top             =   4920
      Width           =   360
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   3
      Left            =   4560
      Picture         =   "frmBattle.frx":BDF0
      Top             =   4920
      Width           =   360
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   2
      Left            =   3360
      Picture         =   "frmBattle.frx":BFEE
      Top             =   4920
      Width           =   360
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   1
      Left            =   2040
      Picture         =   "frmBattle.frx":C226
      Top             =   4920
      Width           =   360
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   0
      Left            =   600
      Picture         =   "frmBattle.frx":C48F
      Top             =   4920
      Width           =   360
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
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
      Left            =   6720
      TabIndex        =   36
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOK 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Left            =   5160
      TabIndex        =   35
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Psynergy:"
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
      Height          =   195
      Left            =   5160
      TabIndex        =   34
      Top             =   2760
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Shape shpList 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblDefend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "De&fend"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   5640
      TabIndex        =   29
      Top             =   5280
      Width           =   1110
   End
   Begin VB.Label lblSummon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Summon"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   4080
      TabIndex        =   28
      Top             =   5280
      Width           =   1365
   End
   Begin VB.Label lblDjinn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Djinn"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3120
      TabIndex        =   27
      Top             =   5280
      Width           =   795
   End
   Begin VB.Label lblPsynergy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Psynergy"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1440
      TabIndex        =   26
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Attack"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   240
      TabIndex        =   25
      Top             =   5280
      Width           =   1050
   End
   Begin VB.Image imgSummon 
      Height          =   1515
      Index           =   7
      Left            =   1680
      Picture         =   "frmBattle.frx":C6D9
      Top             =   1680
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image imgSummon 
      Height          =   1440
      Index           =   6
      Left            =   1680
      Picture         =   "frmBattle.frx":D840
      Top             =   1680
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image imgSummon 
      Height          =   1440
      Index           =   5
      Left            =   1680
      Picture         =   "frmBattle.frx":E953
      Top             =   1680
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image imgSummon 
      Height          =   1410
      Index           =   4
      Left            =   1800
      Picture         =   "frmBattle.frx":F96E
      Top             =   1800
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image imgSummon 
      Height          =   1410
      Index           =   3
      Left            =   1800
      Picture         =   "frmBattle.frx":107AB
      Top             =   1680
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image imgSummon 
      Height          =   465
      Index           =   2
      Left            =   2040
      Picture         =   "frmBattle.frx":114CC
      Top             =   2520
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgSummon 
      Height          =   465
      Index           =   1
      Left            =   1920
      Picture         =   "frmBattle.frx":117D1
      Top             =   2400
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgSummon 
      Height          =   1410
      Index           =   0
      Left            =   1800
      Picture         =   "frmBattle.frx":11D00
      Top             =   1680
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image imgUser 
      Height          =   840
      Index           =   13
      Left            =   720
      Picture         =   "frmBattle.frx":12A21
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgUser 
      Height          =   840
      Index           =   12
      Left            =   1200
      Picture         =   "frmBattle.frx":13050
      Top             =   2280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblHoston 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7680
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Is My Op. Trying To Reset?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   12
      Left            =   7680
      TabIndex        =   22
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image imgDjinn 
      Height          =   825
      Index           =   6
      Left            =   1920
      Picture         =   "frmBattle.frx":136B9
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgDjinn 
      Height          =   825
      Index           =   5
      Left            =   6120
      Picture         =   "frmBattle.frx":1405D
      Top             =   3120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgDjinn 
      Height          =   825
      Index           =   4
      Left            =   1800
      Picture         =   "frmBattle.frx":14839
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgDjinn 
      Height          =   825
      Index           =   3
      Left            =   1800
      Picture         =   "frmBattle.frx":152AC
      Top             =   2160
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgDjinn 
      Height          =   825
      Index           =   2
      Left            =   1920
      Picture         =   "frmBattle.frx":15B4B
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgDjinn 
      Height          =   825
      Index           =   1
      Left            =   1920
      Picture         =   "frmBattle.frx":164C7
      Top             =   2160
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgDjinn 
      Height          =   825
      Index           =   0
      Left            =   1920
      Picture         =   "frmBattle.frx":1700F
      Top             =   2160
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgUser 
      Height          =   1020
      Index           =   11
      Left            =   960
      Picture         =   "frmBattle.frx":17A82
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image imgAttack 
      Height          =   1350
      Left            =   5040
      Picture         =   "frmBattle.frx":18AE9
      Top             =   1320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblHoston 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Is My Op. Waiting?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   7680
      TabIndex        =   20
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblHoston 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7680
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Am I Waiting?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   7680
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Player That Attacks First:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   7
      Left            =   7680
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblHoston 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7680
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgUser 
      Height          =   975
      Index           =   10
      Left            =   960
      Picture         =   "frmBattle.frx":194CF
      Top             =   2280
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgUser 
      Height          =   885
      Index           =   9
      Left            =   600
      Picture         =   "frmBattle.frx":19B9C
      Top             =   1680
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgUser 
      Height          =   885
      Index           =   8
      Left            =   840
      Picture         =   "frmBattle.frx":1A215
      Top             =   2160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Image imgUser 
      Height          =   885
      Index           =   7
      Left            =   240
      Picture         =   "frmBattle.frx":1A865
      Top             =   2040
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgUser 
      Height          =   960
      Index           =   6
      Left            =   240
      Picture         =   "frmBattle.frx":1AEC6
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgUser 
      Height          =   855
      Index           =   5
      Left            =   240
      Picture         =   "frmBattle.frx":1B505
      Top             =   2040
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgUser 
      Height          =   915
      Index           =   4
      Left            =   240
      Picture         =   "frmBattle.frx":1BB64
      Top             =   2040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image imgUser 
      Height          =   900
      Index           =   3
      Left            =   240
      Picture         =   "frmBattle.frx":1C243
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgUser 
      Height          =   900
      Index           =   2
      Left            =   120
      Picture         =   "frmBattle.frx":1C86C
      Top             =   2040
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgUser 
      Height          =   990
      Index           =   1
      Left            =   120
      Picture         =   "frmBattle.frx":1CEB3
      Top             =   1920
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Send a Message:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   6240
      TabIndex        =   14
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblmsg 
      BackStyle       =   0  'Transparent
      Caption         =   "No Messages Recieved"
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
      Height          =   855
      Left            =   6240
      TabIndex        =   11
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   18
      Left            =   5640
      TabIndex        =   10
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Remaining:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   17
      Left            =   5280
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.Line lneDivide 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   352
      X2              =   352
      Y1              =   0
      Y2              =   72
   End
   Begin VB.Image imgYou 
      Height          =   960
      Index           =   3
      Left            =   6000
      MouseIcon       =   "frmBattle.frx":1D54B
      Picture         =   "frmBattle.frx":1DE15
      Top             =   3480
      Width           =   495
   End
   Begin VB.Image imgYou 
      Height          =   915
      Index           =   2
      Left            =   2640
      MouseIcon       =   "frmBattle.frx":1E454
      Picture         =   "frmBattle.frx":1ED1E
      Top             =   3480
      Width           =   525
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   4320
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   1080
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblPP 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "PP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.Image imgYou 
      Height          =   600
      Index           =   1
      Left            =   2760
      Picture         =   "frmBattle.frx":1F3FD
      Top             =   360
      Width           =   345
   End
   Begin VB.Image imgYou 
      Height          =   600
      Index           =   0
      Left            =   120
      Picture         =   "frmBattle.frx":1F7A6
      Top             =   360
      Width           =   345
   End
   Begin VB.Line lneDivide 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   168
      X2              =   168
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Foe's Stats:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Stats:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape shpHP 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   855
   End
   Begin VB.Shape shpHP 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblText 
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
      Height          =   975
      Left            =   840
      TabIndex        =   30
      Top             =   4680
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Shape shpMenu 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   9015
   End
   Begin VB.Image imgYou 
      Height          =   915
      Index           =   7
      Left            =   6750
      MouseIcon       =   "frmBattle.frx":1FB4F
      Picture         =   "frmBattle.frx":20419
      Top             =   3480
      Width           =   525
   End
   Begin VB.Image imgArena 
      Height          =   255
      Left            =   960
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim strWaitAttack As String 'Which attack you're waiting for
Dim intDirMax As Integer 'Maximum number of pictures in a directory
Dim atType As Integer 'critical, special or normal attack
Dim Shake As Integer



Private Sub cmdReset_Click()
'reset battle (debug use)
On Error Resume Next
    If hoston = True Then
        frmHost.Host.SendData "RESET" & vbCrLf
    Else
        frmJoin.Client.SendData "RESET" & vbCrLf
    End If
    
For i = 1 To 4
    bOReady(i) = False
    AttackType(i) = ""
Next 'i
    
'Call EnableChoose

Reset(1) = True
cmdReset.Enabled = False
End Sub

Private Sub cmdSend_Click()
'send a message to other person
On Error GoTo err
If hoston = True Then
frmHost.Host.SendData "GMSG" & txtMsg.Text & vbCrLf
Else
frmJoin.Client.SendData "GMSG" & txtMsg.Text & vbCrLf
End If

If txtMsg.Text = "GETSTATS" Then
    For i = 1 To 2
        If hoston = True Then 'Send Stats
            frmHost.Host.SendData "GETSTATS" & vbCrLf & "HP" & (i + 2) & HP(i) & vbCrLf & "SPEED" & (i + 2) & Speed(i) & vbCrLf & "AP" & (i + 2) & AP(i) & vbCrLf & "DEFENSE" & (i + 2) & Defense(i) & vbCrLf & "PIC" & (i + 2) & Char(i) & vbCrLf & "CHARTYPE" & (i + 2) & CharType(i) & vbCrLf & "RESEARTH" & (i + 2) & intEarthResist(i) & vbCrLf & "RESFIRE" & (i + 2) & intFireResist(i) & vbCrLf & "RESWIND" & (i + 2) & intWindResist(i) & vbCrLf & "RESWATER" & (i + 2) & intWaterResist(i) & vbCrLf & "RESHEART" & (i + 2) & intHeartResist(i) & vbCrLf & "RESDARK" & (i + 2) & intDarkResist(i) & vbCrLf

        Else
            frmJoin.Client.SendData "GETSTATS" & vbCrLf & "HP" & (i + 2) & HP(i) & vbCrLf & "SPEED" & (i + 2) & Speed(i) & vbCrLf & "AP" & (i + 2) & AP(i) & vbCrLf & "DEFENSE" & (i + 2) & Defense(i) & vbCrLf & "PIC" & (i + 2) & Char(i) & vbCrLf & "CHARTYPE" & (i + 2) & CharType(i) & vbCrLf & "RESEARTH" & (i + 2) & intEarthResist(i) & vbCrLf & "RESFIRE" & (i + 2) & intFireResist(i) & vbCrLf & "RESWIND" & (i + 2) & intWindResist(i) & vbCrLf & "RESWATER" & (i + 2) & intWaterResist(i) & vbCrLf & "RESHEART" & (i + 2) & intHeartResist(i) & vbCrLf & "RESDARK" & (i + 2) & intDarkResist(i) & vbCrLf

        End If
    Next 'i
End If

If txtMsg.Text = "exitgame" Then
    Unload Me
    'frmMultiplayer.Show
End If

'If txtMsg.Text <> "root" Then
    txtMsg.Text = ""
'Else
'    Call EnableChoose 'cheater's reset button
'End If





Exit Sub
err:
If GameOver = False Then 'If the game is still going
    Call WriteIni("ERROR", "ERROR", err.Description, App.Path & "\userdata.ini")
    Call WriteIni("ERROR", "SOURCE", "SEND", App.Path & "\userdata.ini")
    Call WriteIni("ERROR", "NUMBER", err.Number, App.Path & "\userdata.ini")
    
    MsgBox "There was an error connecting to your oppontent."
    
    timeBoready.Enabled = False
    
    timePsynergy.Enabled = False
    Unload Me
    frmIntro.Show
Else
    Resume Next
End If

End Sub

Private Sub Form_Activate()
On Error Resume Next
If hoston = True Then
    If frmHost.Host.State = sckClosed Then
        Unload Me
        Exit Sub
    End If
Else
    If frmJoin.Client.State = sckClosed Then
        Unload Me
        Exit Sub
    End If
End If
If BattleLoaded(1) = False Then
    BattleLoaded(1) = True
    'Call DisableChoose 'Don't attack or anything until both players are loaded
    bOReady(1) = False
    If hoston = True Then
        frmHost.Host.SendData "BATTLELOADED" & vbCrLf
    Else
        frmJoin.Client.SendData "BATTLELOADED" & vbCrLf
    End If
End If

If DataSent = False And BattleLoaded(2) = True Then  'If data hasn't been sent yet
    Call LoadBattle
End If





End Sub

Private Sub Form_Load()
On Error Resume Next
timeLoadBattle.Enabled = True
'If BattleLoaded(1) = False Then
'    BattleLoaded(1) = True
'    Call DisableChoose 'Don't attack or anything until both players are loaded
    If hoston = True Then
        frmHost.Host.SendData "BATTLELOADED" & vbCrLf
    Else
        frmJoin.Client.SendData "BATTLELOADED" & vbCrLf
    End If
'End If

If DataSent = False And BattleLoaded(2) = True Then  'If data hasn't been sent yet
    Call LoadBattle
End If



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If lblHP(3).Caption = "Waiting..." Or CharType(4) = "" Then
    For i = 1 To 2
        If hoston = True Then 'Send Stats
            frmHost.Host.SendData "HP" & (i + 2) & HP(i) & vbCrLf & "SPEED" & (i + 2) & Speed(i) & vbCrLf & "AP" & (i + 2) & AP(i) & vbCrLf & "DEFENSE" & (i + 2) & Defense(i) & vbCrLf & "PIC" & (i + 2) & Char(i) & vbCrLf & "CHARTYPE" & (i + 2) & CharType(i) & vbCrLf & "RESEARTH" & (i + 2) & intEarthResist(i) & vbCrLf & "RESFIRE" & (i + 2) & intFireResist(i) & vbCrLf & "RESWIND" & (i + 2) & intWindResist(i) & vbCrLf & "RESWATER" & (i + 2) & intWaterResist(i) & vbCrLf & "RESHEART" & (i + 2) & intHeartResist(i) & vbCrLf & "RESDARK" & (i + 2) & intDarkResist(i) & vbCrLf
            DoEvents
        Else
            frmJoin.Client.SendData "HP" & (i + 2) & HP(i) & vbCrLf & "SPEED" & (i + 2) & Speed(i) & vbCrLf & "AP" & (i + 2) & AP(i) & vbCrLf & "DEFENSE" & (i + 2) & Defense(i) & vbCrLf & "PIC" & (i + 2) & Char(i) & vbCrLf & "CHARTYPE" & (i + 2) & CharType(i) & vbCrLf & "RESEARTH" & (i + 2) & intEarthResist(i) & vbCrLf & "RESFIRE" & (i + 2) & intFireResist(i) & vbCrLf & "RESWIND" & (i + 2) & intWindResist(i) & vbCrLf & "RESWATER" & (i + 2) & intWaterResist(i) & vbCrLf & "RESHEART" & (i + 2) & intHeartResist(i) & vbCrLf & "RESDARK" & (i + 2) & intDarkResist(i) & vbCrLf
            
            DoEvents
        End If
    Next 'i
    If hoston = True Then
        frmHost.Host.SendData "GETSTATS" & vbCrLf
    Else
        frmJoin.Client.SendData "GETSTATS" & vbCrLf
    End If
End If
If lblHP(0).Caption = "Waiting..." Then
    Call LoadBattle
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Sends disconnect, currently not used
On Error Resume Next
DataSent = False
BattleLoaded(1) = False
BattleLoaded(2) = False

If hoston = True Then
    frmHost.Host.SendData "DISC" & vbCrLf
Else
    frmJoin.Client.SendData "DISC" & vbCrLf
End If
frmHost.Host.Close
frmJoin.Client.Close

    frmIntro.Show
    StopMidi

End Sub

Private Sub imgAttack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    imgAttack.Drag
End If
End Sub

Private Sub imgBack_Click()
Call lblBack_Click
End Sub

Private Sub imgIcon_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
    Call lblAttack_Click
ElseIf Index = 1 Then
    Call lblPsynergy_Click
ElseIf Index = 2 Then
    Call lblDjinn_Click
ElseIf Index = 3 Then
    Call lblSummon_Click
ElseIf Index = 4 Then
    Call lblDefend_Click
ElseIf Index = 5 Then
    Call lblStatus_Click
Else
    Call lblBackTurn_Click
End If
End Sub

Private Sub imgOK_Click()
Call lblOK_Click
End Sub

Private Sub imgYou_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
On Error Resume Next
If Index = 1 Then
    MsgBox "You obviously are keen in the areas of battle.  Here is a reward.  Enter 'thereisnocowlevel' without quotes into the textbox in which you choose the arena on the host window to earn a special arena.", vbInformation, "Easter Egg #11"
    Call Encode("11", "EGG11", "EGGL11", App.Path & "\settings.ini")
End If
End Sub

Private Sub imgYou_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If SelectTarget = True And imgYou(Index).MousePointer = 99 Then
    imgYou(2).MousePointer = 0
    imgYou(3).MousePointer = 0
    imgYou(6).MousePointer = 0
    imgYou(7).MousePointer = 0
    SelectTarget = False
    lblSelectTarget.Visible = False
    
    Dim opTarget As Long
    
    If Index = 2 Then
        Target(intTurn) = 1
        opTarget = 3
    ElseIf Index = 3 Then
        Target(intTurn) = 3
        opTarget = 1
    ElseIf Index = 6 Then
        Target(intTurn) = 2
        opTarget = 4
    ElseIf Index = 7 Then
        Target(intTurn) = 4
        opTarget = 2
    End If
    
    If intTurn = 1 Then
        imgTurn.Top = imgYou(4).Top + imgYou(4).Height
        bOReady(1) = True
        If hoston = True Then
            frmHost.Host.SendData "TARGET" & (intTurn + 2) & opTarget & vbCrLf
        Else
            frmJoin.Client.SendData "TARGET" & (intTurn + 2) & opTarget & vbCrLf
        End If
        If strWaitAttack = "ATTACK" Then
            Call SendAttack
        ElseIf strWaitAttack = "PSYNERGY" Then
            Call ComPsynergy
        ElseIf strWaitAttack = "DJINN" Then
            Call ComDjinn
        ElseIf strWaitAttack = "SUMMON" Then
            Call ComSummon
        End If
    Else
        imgTurn.Visible = False
        If hoston = True Then
            frmHost.Host.SendData "TARGET" & (intTurn + 2) & opTarget & vbCrLf
        Else
            frmJoin.Client.SendData "TARGET" & (intTurn + 2) & opTarget & vbCrLf
        End If
        bOReady(2) = True
        If strWaitAttack = "ATTACK" Then
            Call SendAttack
        ElseIf strWaitAttack = "PSYNERGY" Then
            Call ComPsynergy
        ElseIf strWaitAttack = "DJINN" Then
            Call ComDjinn
        ElseIf strWaitAttack = "SUMMON" Then
            Call ComSummon
        End If
    End If
End If
End Sub

Private Sub lblAttack_Click()
'Player Attacks using regular attack
On Error Resume Next
If bAllowAttack = False Or SelectTarget = True Then Exit Sub

strWaitAttack = "ATTACK"
Call subSelectTarget
    
End Sub

Private Sub lblBack_Click()
On Error Resume Next
'closes the psynergy/djinn/summons list
Call HideList
End Sub

Private Sub lblBackTurn_Click()
On Error Resume Next
If intTurn = 2 Then
    intTurn = 1
    bOReady(1) = False
    AttackType(1) = ""
    imgTurn.Top = imgYou(0).Top + imgYou(0).Height
End If

End Sub

Private Sub lblDefend_Click()
On Error Resume Next
If SelectTarget = True Then Exit Sub
'Player defends
If hoston = True Then
    frmHost.Host.SendData "OPREADY" & (intTurn + 2) & vbCrLf
    frmHost.Host.SendData "DEFEND" & (intTurn + 2) & vbCrLf
Else
    frmJoin.Client.SendData "OPREADY" & (intTurn + 2) & vbCrLf
    frmJoin.Client.SendData "DEFEND" & (intTurn + 2) & vbCrLf
End If

AttackType(intTurn) = "DEFEND"

Call DisableChoose

End Sub

Private Sub lblDjinn_Click()
'user wants to use a djinn
On Error Resume Next
If SelectTarget = True Then Exit Sub
lstDjinn.Clear
If RelativeDjinn(1) = 0 Then RelativeDjinn(1) = 1
If intTurn = 1 Then
    For i = 1 To 7
        If modHoverLVL.Djinn(i).Character = intTurn And Djinn(i).Name <> "" Then  'If Djinn exists
            If i <= RelativeDjinn(1) Then
                lstDjinn.AddItem modHoverLVL.Djinn(i).Name
            End If
        End If
    Next 'i
Else
    For i = 8 To 20
        If modHoverLVL.Djinn(i).Character = intTurn And Djinn(i).Name <> "" Then 'If Djinn exists
            If i - 7 <= RelativeDjinn(1) Then
                lstDjinn.AddItem modHoverLVL.Djinn(i).Name
            End If
        End If
    Next 'i
End If


Call ShowList("DJINN")

End Sub

Private Sub lblgen_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If Index = 16 Then
    If Source = lblStatusClose Then
        MsgBox "Did you know that behind the United States, Austraila is country that connects to War of the Adepts the most?", vbInformation, "Easter Egg #23"
        Call Encode("23", "EGG23", "EGGL23", App.Path & "\settings.ini")
    End If
End If

End Sub

Private Sub lblOK_Click()
'user has chosen a summon/psynergy/djinn
On Error Resume Next




'-------Psynergy----------
If lstPsynergy.Visible = True And lstPsynergy.Text <> "" Then
    bOReady(intTurn) = True
    strWaitAttack = "PSYNERGY"
'    Call ComPsynergy
End If


'--------Djinn------------
If lstDjinn.Visible = True And lstDjinn.Text <> "" Then
    bOReady(intTurn) = True
    strWaitAttack = "DJINN"
'    Call ComDjinn
End If 'End Djinn

'--------Summons----------
If lstSummon.Visible = True And lstSummon.Text <> "" Then
    bOReady(intTurn) = True
    strWaitAttack = "SUMMON"
'    Call ComSummon
End If

Call subSelectTarget



If bOReady(1) = True And bOReady(2) = True And bOReady(3) = True And bOReady(4) = True Then  'Are both players ready?
    Call CheckTurns
End If


End Sub

Private Sub lblPsynergy_Click()
'user wants to use a psynergy
On Error Resume Next
If SelectTarget = True Then Exit Sub
Dim iDjinnTotal As Integer

iDjinnTotal = 0
If intTurn = 1 Then
    For i = 1 To 7
        If i <= RelativeDjinn(intTurn) Then
            If Djinn(i).Name <> "" And Djinn(i).Character = intTurn Then
                If Djinn(i).State = 0 Then
                    iDjinnTotal = iDjinnTotal + 1
                End If
            End If
        End If
    Next 'i
Else
    For i = 8 To 20
        If i - 7 <= (RelativeDjinn(1)) Then
            If Djinn(i).Name <> "" And Djinn(i).Character = intTurn Then
                If modHoverLVL.Djinn(i).State = 0 Then
                    iDjinnTotal = iDjinnTotal + 1
                End If
            End If
        End If
    Next 'i
End If



lstPsynergy.Clear
For i = 1 To 30
If modHoverLVL.Psynergy(i).PP <> "" Then
    intCheckPP = CInt(modHoverLVL.Psynergy(i).PP)
Else
    intCheckPP = 999
End If

    
    If modHoverLVL.Psynergy(i).Character = intTurn And modHoverLVL.Psynergy(i).Name <> "" And intCheckPP <= PP(intTurn) Then   'If the Psynergy exists
        If bAllowPsynergy = True Then
            If bAllowHeal = True Or (bAllowHeal = False And modHoverLVL.Psynergy(i).Type <> "HEAL") Then
                If modHoverLVL.Psynergy(i).Djinn <= iDjinnTotal Then
                    lstPsynergy.AddItem strPsyName(i)
                End If
            End If
        End If
    End If
Next 'i


Call ShowList("PSYNERGY")

End Sub

Private Sub lblStatus_Click()
If SelectTarget = True Then Exit Sub
lblStatusClose.Visible = True
txtStatus.Visible = True
txtStatus.Text = ""
txtStatus.Text = "Level After Equalizing: " & RelativeLVL(1)
txtStatus.Text = txtStatus.Text & vbNewLine & CharName(1) & ":"
txtStatus.Text = txtStatus.Text & vbNewLine & "AP: " & AP(1) & " Defense: " & Defense(1)
txtStatus.Text = txtStatus.Text & vbNewLine & "POWER: Earth: " & intEarthPower(1) & " Fire: " & intFirePower(1) & " Wind: " & intWindPower(1) & " Water: " & intWaterPower(1) & " Heart: " & intHeartPower(1) & " Dark: " & intDarkPower(1)
txtStatus.Text = txtStatus.Text & vbNewLine & "RESIST: Earth: " & intEarthResist(1) & " Fire: " & intFireResist(1) & " Wind: " & intWindResist(1) & " Water: " & intWaterResist(1) & " Heart: " & intHeartResist(1) & " Dark: " & intDarkResist(1)
txtStatus.Text = txtStatus.Text & vbNewLine & CharName(2) & ":"
txtStatus.Text = txtStatus.Text & vbNewLine & "AP: " & AP(2) & " Defense: " & Defense(2)
txtStatus.Text = txtStatus.Text & vbNewLine & "POWER: Earth: " & intEarthPower(2) & " Fire: " & intFirePower(2) & " Wind: " & intWindPower(2) & " Water: " & intWaterPower(2) & " Heart: " & intHeartPower(2) & " Dark: " & intDarkPower(2)
txtStatus.Text = txtStatus.Text & vbNewLine & "RESIST: Earth: " & intEarthResist(2) & " Fire: " & intFireResist(2) & " Wind: " & intWindResist(2) & " Water: " & intWaterResist(2) & " Heart: " & intHeartResist(2) & " Dark: " & intDarkResist(2)
txtStatus.Text = txtStatus.Text & vbNewLine & "SPEED1: " & Speed(1) & " SPEED2: " & Speed(2)

End Sub

Private Sub lblStatusClose_Click()
lblStatusClose.Visible = False
txtStatus.Visible = False
End Sub

Private Sub lblStatusClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    lblStatusClose.Drag
End If
End Sub

Private Sub lblSummon_Click()
On Error Resume Next
If SelectTarget = True Then Exit Sub
If bAllowSummon = True Then
    'user wnats to use a psynergy
    lstSummon.Clear
    intDjinnStandby(intTurn) = 0
    If intTurn = 1 Then
        For i = 1 To 7
            If modHoverLVL.Djinn(i).State = 1 And modHoverLVL.Djinn(i).Name <> "" And modHoverLVL.Djinn(i).Character = intTurn And i <= RelativeDjinn(intTurn) Then
                intDjinnStandby(intTurn) = intDjinnStandby(intTurn) + 1
            End If
        Next 'i
    Else
        For i = 8 To 20
            If modHoverLVL.Djinn(i).State = 1 And modHoverLVL.Djinn(i).Name <> "" And modHoverLVL.Djinn(i).Character = intTurn And i - 7 <= RelativeDjinn(intTurn) Then
                intDjinnStandby(intTurn) = intDjinnStandby(intTurn) + 1
            End If
        Next 'i
    End If

    
    For i = 1 To 10
        If modHoverLVL.Summon(i).Name <> "" And modHoverLVL.Summon(i).Character = intTurn Then 'If Summon exists
            If modHoverLVL.Summon(i).Level <= intDjinnStandby(intTurn) Then 'If you have enough Djinn on standby
                lstSummon.AddItem modHoverLVL.Summon(i).Name
            End If ' i <=...
        End If 'strsumname(i)...
    Next 'i
    
    Call ShowList("SUMMON")
Else
    MsgBox "Summons are not allowed in this battle!"
End If

End Sub



Private Sub lstDjinn_Click()
On Error Resume Next
'user selected a djinn
If bDjinnSet(lstDjinn.ListIndex + 1) = 0 Then 'Print Set or Standby
    lblText.Caption = strDjinnName(lstDjinn.ListIndex + 1) & ": " & strDjinnDesc(lstDjinn.ListIndex + 1) & " - {SET}"
ElseIf bDjinnSet(lstDjinn.ListIndex + 1) = 1 Then
    lblText.Caption = strDjinnName(lstDjinn.ListIndex + 1) & ": " & strDjinnDesc(lstDjinn.ListIndex + 1) & " - {STANDBY}"
Else
    lblText.Caption = strDjinnName(lstDjinn.ListIndex + 1) & ": " & strDjinnDesc(lstDjinn.ListIndex + 1) & " - {REST}"
End If

End Sub

Private Sub lstPsynergy_Click()
On Error Resume Next
'user selected a psynergy
Dim intReal As Integer 'Real number of Psynergy
intReal = 1
For i = 1 To 60
    If Psynergy(i).Name <> "" And Psynergy(i).Character = intTurn Then 'if the psynergy exists
        If lstPsynergy.Text = modHoverLVL.Psynergy(i).Name Then
        intReal = i
        End If
    End If
Next 'i

lblText.Caption = modHoverLVL.Psynergy(intReal).Name & ": " & modHoverLVL.Psynergy(intReal).Desc & " - PP: " & modHoverLVL.Psynergy(intReal).PP
End Sub

Private Sub lstSummon_Click()
On Error Resume Next
'user selected a summon

Dim intTempSummon As Long
For i = 1 To 10
    If lstSummon.Text = Summon(i).Name Then
        intTempSummon = i
    End If
Next 'i


lblText.Caption = Summon(intTempSummon).Name & ": " & Summon(intTempSummon).Desc & " - Level: " & Summon(intTempSummon).Level

End Sub

Private Function Damage() As Long
On Error Resume Next
'calculates the damage of a regular attack
Dim itemdamage As Integer
Dim nsave As String
Dim strTime As String
Dim intDamage As Integer
Dim intRand As Long

Randomize
intRand = Int(Rnd * 100) + 1

'ini debug output (more of this should be added
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
nsave = App.Path & "\userdata.ini"

If Luck(intTurn) >= 30 Then Luck(intTurn) = 30
If Luck(intTurn) < 3 Then Luck(intTurn) = 3


'if weapon special goto spcsub
If intRand <= intItemSpcPercent(intWeapon(intTurn)) + Int(Rnd * 3) + 1 Then
    GoTo spcsub
    Exit Function
End If
'if critical goto chsub
If intRand <= Luck(intTurn) + (Int(Rnd * 3) + 1) Then
    GoTo chsub
    Exit Function
End If

' the following is the offical standard damage formula
itemdamage = CInt(strItemDamage(intWeapon(intTurn)))
Call WriteIni("BATTLEDATA", strTime & " First Item Damage", CStr(itemdamage), nsave)
itemdamage = itemdamage / (1.1)
itemdamage = Int((Rnd * (itemdamage * 0.35)) + (itemdamage * 0.2))
Call WriteIni("BATTLEDATA", strTime & " Final Item Damage", CStr(itemdamage), nsave)
Call WriteIni("BATTLEDATA", strTime & " My AP", CStr(AP(1)), nsave)
Call WriteIni("BATTLEDATA", strTime & " Op. Defense", CStr(intoDefense), nsave)

intDamage = AP(intTurn) + itemdamage - Defense(Target(intTurn))
'Comment out when not in ladder tournament:
intDamage = intDamage / 1.45
If intDamage < RelativeLVL(1) Then intDamage = RelativeLVL(1)
intDamage = intDamage + Int(Rnd * 3 + 1) 'add 0 to three random damage
atType = 0
'intDamage = intDamage * 0.9
Damage = intDamage
AttackType(intTurn) = "DAMAGE"

Exit Function

spcsub:
itemdamage = CInt(strItemDamage(intWeapon(intTurn)))
Call WriteIni("BATTLEDATA", strTime & " First Item Damage", CStr(itemdamage), nsave)
itemdamage = itemdamage / (1.1)
itemdamage = Int((Rnd * (itemdamage * 0.35)) + (itemdamage * 0.2) + 1)
Call WriteIni("BATTLEDATA", strTime & " Final Item Damage", CStr(itemdamage), nsave)
Call WriteIni("BATTLEDATA", strTime & " My AP", CStr(AP(1)), nsave)
Call WriteIni("BATTLEDATA", strTime & " Op. Defense", CStr(intoDefense), nsave)

AttackType(intTurn) = "SPATTACK"
intDamage = AP(intTurn) + itemdamage - Defense(Target(intTurn))
If intDamage < 1 Then intDamage = 1
intDamage = intDamage + intItemAddMod(intWeapon(intTurn))
'Commented out because it does too much damge
Dim intSpcPower As Variant
intSpcPower = (1 + GetRelPower(CharType(intTurn)) / 175)
If intSpcPower > 4.5 Then intSpcPower = 4.5
If intSpcPower < 1.5 Then intSpcPower = 1.5

intDamage = intDamage * intSpcPower

intDamage = intDamage + Int(Rnd * 4) 'add 0 to three random damage
'Comment out when not in ladder tournament:
intDamage = intDamage / 1.35

If intDamage <= 0 Then intDamage = 1
atType = 1
intDamage = intDamage
Damage = intDamage

Exit Function

chsub:
itemdamage = CInt(strItemDamage(intWeapon(intTurn)))
Call WriteIni("BATTLEDATA", strTime & " First Item Damage", CStr(itemdamage), nsave)
itemdamage = itemdamage / (1.1)
itemdamage = Int((Rnd * (itemdamage * 0.35)) + (itemdamage * 0.1))
Debug.Print Rnd
Call WriteIni("BATTLEDATA", strTime & " Final Item Damage", CStr(itemdamage), nsave)
Call WriteIni("BATTLEDATA", strTime & " My AP", CStr(AP(1)), nsave)
Call WriteIni("BATTLEDATA", strTime & " Op. Defense", CStr(intoDefense), nsave)

AttackType(intTurn) = "CRATTACK"
intDamage = AP(intTurn) + itemdamage - Defense(Target(intTurn))
If intDamage < 1 Then intDamage = 1
intDamage = intDamage / 1.15
intDamage = intDamage + Int(Rnd * 4) 'add 0 to three random damage

If intDamage <= 0 Then intDamage = 1
atType = 2
intDamage = intDamage
Damage = intDamage

End Function

Private Sub timeBoready_Timer()
On Error Resume Next
'This timer is a 'clean up' timer.  It's purpose is to check
'for errors such as the opponent not being connected anymore
'as well as whether stats haven't been sent, if the game is
'over or if any stray animation is still showing up.

'Exit Sub


'Status Updates:
lblHoston(1).Caption = bOReady(2)  'Player 1 ready?
lblHoston(2).Caption = bOReady(4)  'Player 2 ready?
'If hoston = True Then
'    If intFirstAttack = 1 Then
'        lblHoston(0).Caption = strMyUserName  'Do I host first?
'    Else
'        lblHoston(0).Caption = strOpponent
'    End If
'Else
'    If intFirstAttack = 1 Then
'        lblHoston(0).Caption = strOpponent
'    Else
'        lblHoston(0).Caption = strMyUserName
'    End If
'End If
lblHoston(3).Caption = Reset(2)  'The other player waiting to reset?

'If both players are waiting
If bOReady(2) = True And bOReady(4) = True Then
    Call DisableChoose
    Call CheckTurns
End If

'If any graphics are still showing up
If timeSword.Enabled = False And imgAttack.Visible = True Then
    imgAttack.Visible = False
End If
If timePsynergy.Enabled = False And imgRealPsy.Visible = True Then
    imgRealPsy.Visible = False
End If


'If stats haven't been sent yet
If lblHP(2).Caption = "Waiting..." Then
    If hoston = True Then
        frmHost.Host.SendData "GETSTATS" & vbCrLf
    Else
        frmJoin.Client.SendData "GETSTATS" & vbCrLf
    End If
End If

'If both players hit reset
If Reset(1) = True And Reset(2) = True Then
    EnableChoose
    Reset(1) = False
    Reset(2) = False
End If


'Ping other player
If hoston = True Then
    frmHost.Host.SendData "PING" & vbCrLf
Else
    frmJoin.Client.SendData "PING" & vbCrLf
End If

'Win or Loss:

Exit Sub


If HP(1) <= 0 And HP(2) <= 0 Then 'If I'm dead
    GameOver = True
    If hoston = True Then
        'Comment out for LADDER TOURNAMENT
        'frmChat.Chat.SendData "LADDERLOSS" & vbCrLf
        frmHost.Host.SendData "ILOST" & vbCrLf 'Tell the other player I lost
        DoEvents
        frmHost.Host.Close 'Close the socket
    End If
    If hoston = False Then
            'Comment out for LADDER TOURNAMENT
        'frmChat.Chat.SendData "LADDERLOSS" & vbCrLf
        frmJoin.Client.SendData "ILOST" & vbCrLf
        DoEvents
        frmJoin.Client.Close
    End If
    DidIWin = False 'I did not win
End If
If HP(3) <= 0 And HP(4) <= 0 Then 'If my opponent is dead
    GameOver = True
    If hoston = True Then
            'Comment out for LADDER TOURNAMENT
        'frmChat.Chat.SendData "LADDERWIN" & vbCrLf
        frmHost.Host.SendData "IWIN" & vbCrLf 'Game is over, you can quit.
        DoEvents
        frmHost.Host.Close
    End If
    If hoston = False Then
            'Comment out for LADDER TOURNAMENT
        'frmChat.Chat.SendData "LADDERWIN" & vbCrLf
        frmJoin.Client.SendData "IWIN" & vbCrLf
        DoEvents
        frmJoin.Client.Close
    End If
    DoEvents
    DidIWin = True 'I win
End If

    If HP(1) <= 0 Or HP(2) <= 0 Or HP(3) <= 0 Or HP(4) <= 0 Then 'If either player is dead
        BattleLoaded(1) = False
        BattleLoaded(2) = False
        'Disable timers on this form
        timeSword.Enabled = False
        timePsynergy.Enabled = False
        timeBoready.Enabled = False
        'LADDER TOURNAMENT
        frmEndGame.Show
        If HP(1) > 0 Or HP(2) > 0 Then
            frmChat.Chat.SendData "LADDERWIN" & strMyUserName & vbCrLf
        Else
            frmChat.Chat.SendData "LADDERLOSS" & strMyUserName & vbCrLf
        End If
        
        DataSent = False 'Did not send the stat increase to the server yet
        Unload Me
    End If

Exit Sub
err:

If GameOver = False Then 'If the game is still going on
    MsgBox "There was an error connecting to your opponent. (" & err.Description & ")"
    'Close sockets
    If hoston = True Then
    frmHost.Host.Close
    Else
    frmJoin.Client.Close
    End If
    'Disable timers on this form
    timeSword.Enabled = False
    timePsynergy.Enabled = False
    timeBoready.Enabled = False
    
    frmMultiplayer.Show
    
    DataSent = False 'Did not send the stat increase to the server yet
    Unload Me
Else
    Resume Next
End If

End Sub

Private Sub timeCount_Timer()
'Not currently used

' selection timer
'On Error Resume Next
'curCount = curCount - 1 'Reduce time left
'lblGen(18).Caption = curCount 'Display time left

'    If curCount = 0 Then 'If you're out of time
'        If bOReady(1) = False Then
'            frmHost.Host.SendData "BOREADY" & vbCrLf
'            frmHost.Host.SendData "DEFEND" & vbCrLf 'Auto-defend
'            bOReady(1) = True
'            bWaitDefend(1) = True
'            timecount.Enabled = False
'            curCount = 20 'Reset value
'            lblGen(18).Caption = "20"
'            Call DisableChoose
'        End If 'boready(1) = false
'    End If 'curcount = 0
    
End Sub

Private Sub timeDjinn_Timer()
' djinn graphics timer
On Error Resume Next

If imgDjinn(0).Visible = False Then imgDjinn(0).Visible = True

If CurrentOp = 1 Then
    imgDjinn(0).Left = imgUser(2).Left - 15 'Djinn comes down on Player 1
End If
If CurrentOp = 2 Then
    imgDjinn(0).Left = imgUser(3).Left - 15 'Djinn comes down on Player 2
End If

imgDjinn(0).Top = imgDjinn(0).Top + 5

If imgDjinn(0).Top <= imgUser(2).Top - 15 Then
    timeDjinn.Enabled = False
    imgDjinn(0).Visible = False
    imgDjinn(0).Top = 120
End If

End Sub


Private Function DoPsy(currentPsy As Integer) As Long
On Error Resume Next
'calculates psynergy attack dammage
Dim PsyDamage As Integer
Dim PowerMult As Variant  'Relative Multiplyer

Dim nsave As String
Dim strTime As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
nsave = App.Path & "\userdata.ini"


PsyDamage = CInt(modHoverLVL.Psynergy(currentPsy).Damage)
Call WriteIni("PSYNERGY", strTime & "1stDAMAGE", CStr(PsyDamage), nsave)

PowerMult = (GetRelPower(CharType(1)) / 200)
Call WriteIni("PSYNERGY", strTime & "RelPowerDiv200", CStr(PowerMult), nsave)
PowerMult = PowerMult + 1
Call WriteIni("PSYNERGY", strTime & "POWMULT", CStr(PowerMult), nsave)
PsyDamage = PowerMult * PsyDamage 'Base damage * relative power
Call WriteIni("PSYNERGY", strTime & "2ndDAMAGE", CStr(PsyDamage), nsave)
PsyDamage = PsyDamage + (Int(Rnd * 3) + 1) 'Add 1 to 3 random damage
Call WriteIni("PSYNERGY", strTime & "RandDAMAGE", CStr(PsyDamage), nsave)

If PsyDamage < 1 Then PsyDamage = 1 'If the attack does less than 1 damage

'Comment out when not in ladder tournament:
PsyDamage = PsyDamage / 1.45

DoPsy = PsyDamage 'return the damage done

End Function
Private Function StatIncrease(curPsy) As Long
On Error Resume Next
'stat increase dammage calculations
Dim PsyDamage As Integer
PsyDamage = CInt(strPsyDamage(curPsy))
PsyDamage = PsyDamage - (PsyDamage * 0.1)

If PsyDamage < 1 Then
StatIncrease = 1
Else
StatIncrease = PsyDamage
End If

End Function


Private Sub timeLoadBattle_Timer()
On Error Resume Next
Exit Sub

If lblHP(2).Caption = "Waiting..." Then
    BattleLoaded(1) = True
    'Call DisableChoose 'Don't attack or anything until both players are loaded
    If hoston = True Then
        frmHost.Host.SendData "GETSTATS" & vbCrLf
    Else
        frmJoin.Client.SendData "GETSTATS" & vbCrLf
    End If
    timeLoadBattle.Enabled = False
    Call LoadBattle
End If


        
'If lblgen(10).Caption = "Waiting..." Or lblHoston(0).Caption = "." Then
'    If BattleLoaded(1) = False Or BattleLoaded(2) = False Then
'        BattleLoaded(1) = True
'        Call DisableChoose 'Don't attack or anything until both players are loaded
'        If hoston = True Then
'            frmHost.Host.SendData "BATTLELOADED" & vbCrLf
'        Else
'            frmJoin.Client.SendData "BATTLELOADED" & vbCrLf
'        End If
'    End If
        
'    If HP(1) <= 0 Then
'        Call LoadBattle
'    End If
    
'    If DataSent = False And BattleLoaded(2) = True Then  'If data hasn't been sent yet
'        Call LoadBattle
'    End If
'End If
End Sub

Private Sub timePsynergy_Timer()
'psynergy animation timer
On Error Resume Next

'If CurrentOp = 1 Then
'    imgRealPsy.Left = imgYou(2).Left - 3
'End If
'If CurrentOp = 2 Then
'    imgRealPsy.Left = imgYou(3).Left + 3
'End If

If imgRealPsy.Visible = False Then
    imgRealPsy.Visible = True
    strCurPsyDir = GetDirName(strCurPsyDir)
    'Number of frames
    intDirMax = CInt(GetFromIni("GEN", strCurPsyDir, App.Path & "\animation.ini"))
    PsyFrame = 0
End If

PsyFrame = PsyFrame + 1

Dim strPsyFrame As String
If PsyFrame < 10 Then
    strPsyFrame = "00" & CStr(PsyFrame)
ElseIf PsyFrame < 100 Then
    strPsyFrame = "0" & CStr(PsyFrame)
Else
    strPsyFrame = CStr(PsyFrame)
End If

If CurrentOp = 1 Then
    strPsyFrame = strPsyFrame & ".gif"
Else
    strPsyFrame = strPsyFrame & "f.gif" 'Forward
End If

imgRealPsy.Picture = LoadPicture(App.Path & "\" & strCurPsyDir & "\" & strPsyFrame)

If PsyFrame >= intDirMax Then
    timePsynergy.Enabled = False
    imgRealPsy.Visible = False
    PsyFrame = 0
End If

End Sub

Private Sub timeSummon_Timer()
'summon animation timer
On Error Resume Next

If imgSummon(0).Visible = False Then
    If CurrentOp = 1 Then
        imgSummon(0).Left = 0 - imgSummon(0).Width 'Left boundary
    Else
        imgSummon(0).Left = Me.ScaleWidth 'Right boundary
    End If
    imgSummon(0).Visible = True
End If
    
If CurrentOp = 1 Then

    imgSummon(0).Left = imgSummon(0).Left + 6
    If imgSummon(0).Left >= Me.ScaleWidth Then
        timeSummon.Enabled = False
        imgSummon(0).Visible = False
    End If
    
End If

If CurrentOp = 2 Then
    imgSummon(0).Left = imgSummon(0).Left - 6
    
    If imgSummon(0).Left <= 0 Then
        timeSummon.Enabled = False
        imgSummon(0).Visible = False
    End If
    
End If


Shake = Shake + 1 'Variable for "shaking" the screen

If Shake = 1 Then 'Move the screen left and up
    frmBattle.Top = frmBattle.Top - 25
    frmBattle.Left = frmBattle.Left - 25
End If
If Shake = 2 Then 'Move the screen right and down
    frmBattle.Top = frmBattle.Top + 25
    frmBattle.Left = frmBattle.Left + 25
End If

If Shake = 3 Then Shake = 0 'Reset variable

End Sub

Private Sub timeSword_Timer()
'sword animation timer
On Error Resume Next
If imgAttack.Visible = False Then imgAttack.Visible = True

'Set X coord for sword
If CurrentOp = 1 Then
    imgAttack.Left = 168
End If
If CurrentOp = 2 Then
    imgAttack.Left = 344
End If

imgAttack.Top = imgAttack.Top + 2.5 'Make sword come down

If imgAttack.Top >= imgUser(2).Top - 12 Then
    imgAttack.Top = 88 'Reset sword's position
    timeSword.Enabled = False
    imgAttack.Visible = False
End If

End Sub


Private Sub timeWait_Timer()
On Error Resume Next
'This timer pauses between each player's attack to
'give them time to read the message.
'It will also load the end game form if a player has lost



timeWait.Enabled = False
Call DoAttacks
'Call CheckTurns



'Win or Loss:
If BattleLoaded(1) = True And BattleLoaded(2) = True Then
    If HP(1) <= 0 And HP(2) <= 0 Then 'If I'm dead
        GameOver = True
        If hoston = True Then
                'Comment out for LADDER TOURNAMENT
            'frmChat.Chat.SendData "LADDERLOSS" & vbCrLf
            frmHost.Host.SendData "ILOST" & vbCrLf 'Tell the other player I lost
            DoEvents
            frmHost.Host.Close 'Close the socket
        End If
        If hoston = False Then
                'Comment out for LADDER TOURNAMENT
            'frmChat.Chat.SendData "LADDERLOSS" & vbCrLf
            frmJoin.Client.SendData "ILOST" & vbCrLf
            DoEvents
            frmJoin.Client.Close
        End If
        DidIWin = False 'I did not win
    End If
    If HP(3) <= 0 And HP(4) <= 0 Then 'If my opponent is dead
        GameOver = True
        If hoston = True Then
                'Comment out for LADDER TOURNAMENT
            frmChat.Chat.SendData "LADDERWIN" & vbCrLf
            frmHost.Host.SendData "IWIN" & vbCrLf 'Game is over, you can quit.
            DoEvents
            frmHost.Host.Close
        End If
        If hoston = False Then
                'Comment out for LADDER TOURNAMENT
            frmChat.Chat.SendData "LADDERWIN" & vbCrLf
            frmJoin.Client.SendData "IWIN" & vbCrLf
            DoEvents
            frmJoin.Client.Close
        End If
        DoEvents
        DidIWin = True 'I win
    End If
    
    If (HP(1) <= 0 And HP(2) <= 0) Or (HP(3) <= 0 And HP(4) <= 0) Then 'If either player is dead
        BattleLoaded(1) = False
        BattleLoaded(2) = False
        'Disable timers on this form
        timeSword.Enabled = False
        timePsynergy.Enabled = False
        timeBoready.Enabled = False
        timeWait.Enabled = False
        'LADDER TOURNAMENT
        frmEndGame.Show
        If HP(1) > 0 Or HP(2) > 0 Then
        frmChat.Chat.SendData "LADDERWIN" & strMyUserName & vbCrLf
        Else
        frmChat.Chat.SendData "LADDERLOSS" & strMyUserName & vbCrLf
        End If
        DataSent = False 'Did not send the stat increase to the server yet
        Unload Me
    End If
End If

Call AutoScrollTxt(txtDisplay)


End Sub

Private Sub txtDisplay_Change()
On Error Resume Next
Call AutoScrollTxt(txtDisplay)
End Sub

Private Sub txtDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeySpace Then
    MsgBox "Obviously you're very interested in 'space'.  Maybe you should start watching a very funny animated show on Fox from 7-7:30 PM on Sundays (when there isn't football) called Futurama.  If you want to find out more, take a look at the website that will pop up.", vbInformation, "Easter Egg #20"
    frmBrowser.Show
    frmBrowser.Web.Navigate "http://www.gotfuturama.com"
    Call Encode("20", "EGG20", "EGGL20", App.Path & "\settings.ini")
End If
End Sub

Private Sub txtMsg_Change()
On Error Resume Next
If txtMsg.Text = "my cat's breath smells like cat food" Then
    MsgBox "I bent my wookie.", vbInformation, "Easter Egg #13"
    Call Encode("13", "EGG13", "EGGL13", App.Path & "\settings.ini")
End If
If Right$(txtMsg.Text, 1) = vbCrLf Then
    txtMsg.Text = Left$(txtMsg.Text, Len(txtMsg.Text) - 1)
End If
If txtMsg.SelText = "egg26" And txtMsg.Text <> txtMsg.SelText Then
    If CurEgg26 = 2 Then
        CurEgg26 = 3
        Call PlySound("explosion")
    Else
        CurEgg26 = 1
    End If
End If
End Sub

Private Sub txtmsg_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'txtmsg keydown stuff
If keyascii = 13 Then keyascii = 0
If KeyCode = vbKeyReturn Then
    Call cmdSend_Click
End If


End Sub
Private Function DjinnDo(curDjinn) As Long
On Error Resume Next
Randomize
'djinn dammage calculator
Dim DjinnDamage As Integer
DjinnDamage = Djinn(curDjinn).Damage
DjinnDamage = Int((Rnd * (DjinnDamage * 0.25) + (DjinnDamage * 0.05)) + 1)
If DjinnDamage < 1 Then
    DjinnDamage = 1
End If
DjinnDo = DjinnDamage
End Function
Private Function SummonDamage(ByVal curSummon As Integer) As Long
On Error Resume Next
'Official Summon Damage Formula
'Random 1 to 3 + 30/60/120/240 summon lvl + 0.03/0.06/0.09/0.12 multiplier of Max HP
'All of this times relative power

Dim strTime As String
Dim nsave As String

strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
nsave = App.Path & "\userdata.ini"


Dim RandInt As Integer
Randomize
RandInt = Int(Rnd * 3) + 1
Call WriteIni("SUMMON", strTime & "RANDDMG", CStr(RandInt), nsave)
If curSummon = 1 Then
    RandInt = RandInt + 30
    Call WriteCharType("POWER", 1, 10)
    intSumBoost = 10
ElseIf curSummon = 2 Then
    RandInt = RandInt + 60
    Call WriteCharType("POWER", 1, 30)
    intSumBoost = 30
ElseIf curSummon = 3 Then
    RandInt = RandInt + 120
    Call WriteCharType("POWER", 1, 60)
    intSumBoost = 60
Else
    RandInt = RandInt + 240
    Call WriteCharType("POWER", 1, 100)
    intSumBoost = 100
End If

intSumBoost = intSumBoost / 2

intEarthPower(1) = intEarthPower(1) + intSumBoost
intFirePower(1) = intFirePower(1) + intSumBoost
intWindPower(1) = intWindPower(1) + intSumBoost
intWaterPower(1) = intWaterPower(1) + intSumBoost
intHeartPower(1) = intHeartPower(1) + intSumBoost
intDarkPower(1) = intDarkPower(1) + intSumBoost


Call WriteIni("SUMMON", strTime & "BASEDMG", CStr(RandInt), nsave)

RandInt = RandInt / 1.5 'Not official, but lessens the damage because Summons currently do too much damage

RandInt = RandInt + (MaxHP(Target(intTurn)) * (0.03 * curSummon))
Call WriteIni("SUMMON", strTime & "MAXHP", CStr(MaxHP(2)), nsave)
Call WriteIni("SUMMON", strTime & "MAXHPMULT", CStr(RandInt), nsave)

Dim varElementalMult As Variant
varElementalMult = (1 + (GetRelPower(CharType(1)) / 200))

Call WriteIni("SUMMON", strTime & "TYPEMULT", CStr(varElementalMult), nsave)

RandInt = RandInt * varElementalMult 'Times it by Psynergy multiplier

If RandInt < 20 Then RandInt = 20 'Minimum summon damage

Call WriteIni("SUMMON", strTime & "FINALDMG", CStr(RandInt), nsave)

'Comment out for LADDER TOURNAMENT
RandInt = RandInt / 2.1

SummonDamage = RandInt

End Function
Sub DisableChoose()
On Error Resume Next
'Gets rid of menu options

If intTurn < 3 Then

    Call HideList
    If intTurn = 1 Then
        imgTurn.Top = imgYou(4).Top + imgYou(4).Height
        bOReady(1) = True
        If HP(2) <= 0 Then
            bOReady(2) = True
            AttackType(2) = "DEAD"
            intTurn = 3
            Call DisableChoose
            Exit Sub
        Else
            intTurn = intTurn + 1
        End If
    Else
        imgTurn.Visible = False
        bOReady(2) = True
        intTurn = intTurn + 1
        Call DisableChoose
        Exit Sub
    End If
    
Else

    imgTurn.Visible = False
    timecount.Enabled = False
    lblAttack.Visible = False
    lblPsynergy.Visible = False
    lblDjinn.Visible = False
    lblSummon.Visible = False
    lblDefend.Visible = False
    shpMenu.Visible = False
    lstPsynergy.Visible = False
    lstSummon.Visible = False
    lstDjinn.Visible = False
    shpList.Visible = False
    lblOK.Visible = False
    lblDesc.Visible = False
    lblBack.Visible = False
    lblBackTurn.Visible = False
    'lblText.Visible = True
    lblText.Caption = ""
    imgOK.Visible = False
    imgBack.Visible = False
    PlayerWait = 0
    lblStatusClose.Visible = False
    txtStatus.Visible = False
    lblStatus.Visible = False
    
    For i = 0 To 6
        frmBattle.imgIcon(i).Visible = False
    Next 'i
    
    
    If bOReady(4) = True Then
        Call DoAttacks
    End If
    
End If

End Sub

Sub ShowList(ByVal strData As String)
On Error Resume Next
'Shows list choosing for Psynergy, Djinn, Summon

shpList.Visible = True
imgOK.Visible = True
imgBack.Visible = True
lblText.Caption = ""
lblText.Visible = True
shpMenu.Visible = False
lblAttack.Visible = False
lblPsynergy.Visible = False
lblDjinn.Visible = False
lblSummon.Visible = False
lblDefend.Visible = False
lblStatus.Visible = False
For i = 0 To 5
    imgIcon(i).Visible = False
Next 'i

If strData = "PSYNERGY" Then
    lstPsynergy.Visible = True
    lstDjinn.Visible = False
    lstSummon.Visible = False
    lblDesc.Caption = "Select Psynergy:"
End If
If strData = "DJINN" Then
    lstPsynergy.Visible = False
    lstDjinn.Visible = True
    lstSummon.Visible = False
    lblDesc.Caption = "Select Djinn:"
End If
If strData = "SUMMON" Then
    lstPsynergy.Visible = False
    lstDjinn.Visible = False
    lstSummon.Visible = True
    lblDesc.Caption = "Select Summon:"
End If

End Sub
Sub HideList()
On Error Resume Next
'Shows menus, hides list

lblAttack.Visible = True
lblPsynergy.Visible = True
lblDjinn.Visible = True
lblSummon.Visible = True
lblDefend.Visible = True
shpMenu.Visible = True
lstSummon.Visible = False
lstPsynergy.Visible = False
lstDjinn.Visible = False
lblText.Caption = ""
imgOK.Visible = False
lblDesc.Visible = False
imgBack.Visible = False
lblText.Visible = False
shpList.Visible = False
lblStatus.Visible = True

For i = 0 To 5
    imgIcon(i).Visible = True
Next 'i
End Sub

Function GetCharType(ByVal PowOrResist As String, ByVal Player As Integer, ByVal strTypo As String) As Integer
On Error Resume Next
'gets power or resist of a based on a given type
If Player = 1 Then
    If PowOrResist = "POWER" Then
        Select Case strTypo
        Case "E"
            GetCharType = intEarthPower(1)
        Case "F"
            GetCharType = intFirePower(1)
        Case "N"
            GetCharType = intWindPower(1)
        Case "W"
            GetCharType = intWaterPower(1)
        Case "H"
            GetCharType = intHeartPower(1)
        Case "D"
            GetCharType = intDarkPower(1)
        End Select
    Else
        Select Case strTypo
        Case "E"
            GetCharType = intEarthResist(1)
        Case "F"
            GetCharType = intFireResist(1)
        Case "N"
            GetCharType = intWindResist(1)
        Case "W"
            GetCharType = intWaterResist(1)
        Case "H"
            GetCharType = intHeartResist(1)
        Case "D"
            GetCharType = intDarkResist(1)
        End Select
    End If
Else
    If PowOrResist = "POWER" Then
        Select Case strTypo
        Case "E"
            GetCharType = intEarthPower(2)
        Case "F"
            GetCharType = intFirePower(2)
        Case "N"
            GetCharType = intWindPower(2)
        Case "W"
            GetCharType = intWaterPower(2)
        Case "H"
            GetCharType = intHeartPower(2)
        Case "D"
            GetCharType = intDarkPower(2)
        End Select
    Else
        Select Case strTypo
        Case "E"
            GetCharType = intEarthResist(2)
        Case "F"
            GetCharType = intFireResist(2)
        Case "N"
            GetCharType = intWindResist(2)
        Case "W"
            GetCharType = intWaterResist(2)
        Case "H"
            GetCharType = intHeartResist(2)
        Case "D"
            GetCharType = intDarkResist(2)
        End Select
    End If
End If
End Function


Public Function GetRelPower(strTypo As String) As Variant
On Error Resume Next
'Gets the relative (power - resist) power of two characters
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
Dim nsave As String
nsave = App.Path & "\userdata.ini"
Call WriteIni("PSYNERGY", strTime & "PsyType", strTypo, nsave)

Dim CurPower As Integer, CurResist As Integer
CurPower = GetCharType("POWER", 1, strTypo) 'Get my power
Call WriteIni("PSYNERGY", strTime & "CURPOWER", CStr(CurPower), nsave)
CurResist = GetCharType("RESIST", 2, strTypo) 'Get my opponent's resist
Call WriteIni("PSYNERGY", strTime & "CURRESIST", CStr(CurResist), nsave)
Dim currelpower As Integer
currelpower = CurPower - CurResist
'If currelpower > 75 Then currelpower = 75


GetRelPower = currelpower
End Function
Public Sub ComPsynergy()
On Error Resume Next

    Dim stretime As String 'Current time
    Dim ssave As String 'Save to userdata.ini
    stretime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
    ssave = App.Path & "\userdata.ini"
    
    
    If hoston = True Then
        frmHost.Host.SendData "OPREADY" & (intTurn + 2) & vbCrLf 'I'm ready
    Else
        frmJoin.Client.SendData "OPREADY" & (intTurn + 2) & vbCrLf 'I'm ready
    End If
    

    Dim strPsyCheck As String
    strPsyCheck = lstPsynergy.Text
    For i = 1 To 60
        If strPsyCheck = modHoverLVL.Psynergy(i).Name Then 'Do the names match up?
            RealPsy = i
        End If
    Next 'i
    
    PP(intTurn) = PP(intTurn) - CInt(modHoverLVL.Psynergy(RealPsy).PP)
    
    '---------Do Attacking Psynergy------
    If modHoverLVL.Psynergy(RealPsy).Type = "DAMAGE" Then 'If it's an attacking Psynergy
        AttackDamage(intTurn) = DoPsy(RealPsy) 'Get damage value
        DoEvents
    
        AttackType(intTurn) = "PSY"
    
        If hoston = True Then
        frmHost.Host.SendData "DOPSY" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf 'I'm ready to attack w/ Psynergy
        End If
        If hoston = False Then
        frmJoin.Client.SendData "DOPSY" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf 'I'm ready to attack w/ Psynergy
        End If
    
    
    ElseIf modHoverLVL.Psynergy(RealPsy).Type <> "HEAL" Then 'Not an attack Psynergy and not a heal Psynergy
    
        '-------Do Stat Increases-----------
        AttackDamage(intTurn) = StatIncrease(RealPsy) 'Get stat increased
    
    End If
    
    
    '---------Do Heal-------------
    If modHoverLVL.Psynergy(RealPsy).Type = "HEAL" Then
        bWaitHeal(1) = True 'Waiting to heal
        
        'Healing Psynergy will always heal the same amount of damage
        AttackDamage(intTurn) = CInt(strPsyDamage(RealPsy))
        
        If HP(1) + AttackDamage(intTurn) > MaxHP(1) Then 'Can't heal more than max
            AttackDamage(intTurn) = AttackDamage(intTurn) - ((HP(1) + AttackDamage(intTurn)) - MaxHP(1))
        End If
        
        If hoston = True Then
        frmHost.Host.SendData "HEAL" & (intTurn + 2) & vbCrLf
        frmHost.Host.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf 'Increase HP by this much
        End If
        If hoston = False Then
        frmJoin.Client.SendData "HEAL" & (intTurn + 2) & vbCrLf
        frmJoin.Client.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
    End If
    
    '-------------Do Boost PP-------------
    If modHoverLVL.Psynergy(RealPsy).Type = "PP" Then
        bWaitBoostPP(1) = True
        
        If hoston = True Then
        frmHost.Host.SendData "BOOSTPP" & (intTurn + 2) & vbCrLf
        frmHost.Host.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
        If hoston = False Then
        frmJoin.Client.SendData "BOOSTPP" & (intTurn + 2) & vbCrLf
        frmJoin.Client.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
        
    End If
    
    '------------Do Boost Attack------------
    If modHoverLVL.Psynergy(RealPsy).Type = "ATTACK" Then
        bWaitBoostAttack(1) = True
        
        If hoston = True Then
        frmHost.Host.SendData "BOOSTAP" & (intTurn + 2) & vbCrLf
        frmHost.Host.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
        If hoston = False Then
        frmJoin.Client.SendData "BOOSTAP" & (intTurn + 2) & vbCrLf
        frmJoin.Client.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
    End If
    
    '-------------Do Boost Defense------------
    If modHoverLVL.Psynergy(RealPsy).Type = "DEFENSE" Then
        bWaitBoostDefense(1) = True
        
        If hoston = True Then
        frmHost.Host.SendData "BOOSTDEFENSE" & (intTurn + 2) & vbCrLf
        frmHost.Host.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
        If hoston = False Then
        frmJoin.Client.SendData "BOOSTDEFENSE" & (intTurn + 2) & vbCrLf
        frmJoin.Client.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
    End If
    
    '---------------Do Drop Attack----------
    If modHoverLVL.Psynergy(RealPsy).Type = "DROPATTACK" Then
        bWaitDropAttack(1) = True
        
        If hoston = True Then
        frmHost.Host.SendData "DROPATTACK" & (intTurn + 2) & vbCrLf
        frmHost.Host.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
        If hoston = False Then
        frmJoin.Client.SendData "DROPATTACK" & (intTurn + 2) & vbCrLf
        frmJoin.Client.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
    End If
    
    '----------Do Drop Defense---------------
    If modHoverLVL.Psynergy(RealPsy).Type = "DROPDEFENSE" Then
        bWaitDropDefense(1) = True
        
        If hoston = True Then
        frmHost.Host.SendData "DROPDEFENSE" & (intTurn + 2) & vbCrLf
        frmHost.Host.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
        If hoston = False Then
        frmJoin.Client.SendData "DROPDEFENSE" & (intTurn + 2) & vbCrLf
        frmJoin.Client.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        End If
    End If

    Call DisableChoose
    
End Sub
Public Sub ComDjinn()
On Error Resume Next
Dim intDjinn As Long
For i = 1 To 20
    If lstDjinn.Text = Djinn(i).Name Then
        intDjinn = i
    End If
Next 'i

    If Djinn(intDjinn).State = 0 Then 'If the Djinn is set
        Djinn(intDjinn).State = 1 'unset djinn
        thedamage(1) = DjinnDo(lstDjinn.ListIndex + 1) 'Get damage for Djinn
        intDjinnStandby(intTurn) = intDjinnStandby(intTurn) + 1 'total number of standby djinn increase

        
        Dim strDjinnSend As String
        'Set the Djinn picture depending on Elemental Type of Player
        If CharType(1) = "W" Then strDjinnSend = "4"
        If CharType(1) = "E" Then strDjinnSend = "1"
        If CharType(1) = "F" Then strDjinnSend = "2"
        If CharType(1) = "N" Then strDjinnSend = "3"
        If CharType(1) = "D" Then strDjinnSend = "5"
        If CharType(1) = "H" Then strDjinnSend = "6"
        
        Select Case strDjinnType(lstDjinn.ListIndex + 1)


        Case "DAMAGE"
            AttackType(intTurn) = "DJINNDAMAGE"
            

            'Dim dam As Variant
            Dim intRelPower As Variant
            'thedamage(1) = CInt(strDjinnDamage(lstDjinn.ListIndex + 1))
            'thedamage(1) = thedamage(1) * 0.25
            'intRelPower = (GetRelPower(CharType(1)) / 10)
            'If intRelPower < 0.75 Then intRelPower = 0.75
            'If intRelPower > 1.95 Then intRelPower = 1.95
            'dam = ((AP(1) - intoDefense + thedamage(1))) * intRelPower
            'dam = dam + Rnd(4)
            'If dam < 1 Then dam = 1
            'thedamage(1) = dam
            
            thedamage(1) = CInt(strItemDamage(intWeapon(intTurn))) * 0.75
            'Call WriteIni("BATTLEDATA", strTime & " First Djinn Damage", CStr(thedamage(1)), nSave)
            thedamage(1) = thedamage(1) / (1.1)
            thedamage(1) = Int((Rnd * (thedamage(1) * 0.35)) + (thedamage(1) * 0.2))
            'Call WriteIni("BATTLEDATA", strTime & " Final Djinn Damage", CStr(thedamage(1)), nSave)
            'Call WriteIni("BATTLEDATA", strTime & " My AP", CStr(AP(1)), nSave)
            'Call WriteIni("BATTLEDATA", strTime & " Op. Defense", CStr(intoDefense), nSave)
            
            thedamage(1) = AP(intTurn) + thedamage(1) - Defense(Target(intTurn))
            'intDamage = intDamage / 2
            If thedamage(1) < RelativeLVL(1) Then thedamage(1) = RelativeLVL(1)
            thedamage(1) = thedamage(1) + Int(Rnd * 3 + 1) 'add 0 to three random damage
            intRelPower = (GetRelPower(CharType(intTurn)) / 10)
            If intRelPower < 0.75 Then intRelPower = 0.75
            If intRelPower > 1.75 Then intRelPower = 1.75
            thedamage(1) = thedamage(1) * intRelPower
            If thedamage(1) < 1 Then thedamage(1) = 1
            thedamage(1) = thedamage(1) / 2
            
            
            AttackDamage(intTurn) = thedamage(1)


            If hoston = True Then
                frmHost.Host.SendData "DJINNDAMAGE" & (intTurn + 2) & vbCrLf
                frmHost.Host.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                frmHost.Host.SendData "DJINNTYPE" & (intTurn + 2) & strDjinnSend & vbCrLf
                DoEvents
            End If
            If hoston = False Then
                frmJoin.Client.SendData "DJINNDAMAGE" & (intTurn + 2) & vbCrLf
                frmJoin.Client.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                frmJoin.Client.SendData "DJINNTYPE" & (intTurn + 2) & strDjinnSend & vbCrLf
                DoEvents
            End If
        
        Case "HEAL" 'If the Djinn Heals
            AttackType(intTurn) = "DJINNHEAL"
            If bAllowHeal = True Then
                thedamage(1) = CInt(Djinn(intDjinn).Damage)
                thedamage(1) = thedamage(1) / 2
                'If HP(1) + thedamage(1) >= MaxHP(1) Then
                '    Dim intTempDamage As Integer
                '    intTempDamage = thedamage(1) + HP(1)
                '    intTempDamage = intTempDamage - MaxHP(1)
                '    thedamage(1) = thedamage(1) - intTempDamage
                'End If
                AttackDamage(intTurn) = thedamage(1)
            Else
                AttackDamage(intTurn) = 0
            End If
            
            If hoston = True Then
                frmHost.Host.SendData "DJINNHEAL" & (intTurn + 2) & vbCrLf
                frmHost.Host.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
            If hoston = False Then
                frmJoin.Client.SendData "DJINNHEAL" & (intTurn + 2) & vbCrLf
                frmJoin.Client.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
        
        
        Case "PP" 'If the Djinn raises PP
            AttackType(intTurn) = "DJINNBOOSTPP"
            AttackDamage(intTurn) = thedamage(1)
            If hoston = True Then
                frmHost.Host.SendData "DJINNPP" & (intTurn + 2) & vbCrLf
                frmHost.Host.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
            If hoston = False Then
                frmJoin.Client.SendData "DJINNPP" & (intTurn + 2) & vbCrLf
                frmJoin.Client.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
        
        
        Case "DROPATTACK" 'If the Djinn drops enemy attack
            AttackType(intTurn) = "DJINNREDUCEAP"
            AttackDamage(intTurn) = thedamage(1)
            If hoston = True Then
                frmHost.Host.SendData "DJINNDROPATTACK" & (intTurn + 2) & vbCrLf
                frmHost.Host.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
            If hoston = False Then
                frmJoin.Client.SendData "DJINNDROPATTACK" & (intTurn + 2) & vbCrLf
                frmJoin.Client.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
                
        Case "DROPDEFENSE" 'If the Djinn drops enemy defense
            AttackType(intTurn) = "DJINNREDUCEDEFENSE"
            AttackDamage(intTurn) = thedamage(1)
            If hoston = True Then
                frmHost.Host.SendData "DJINNDROPDEFENSE" & (intTurn + 2) & vbCrLf
                frmHost.Host.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
            If hoston = False Then
                frmJoin.Client.SendData "DJINNDROPDEFENSE" & (intTurn + 2) & vbCrLf
                frmJoin.Client.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
            
        Case "DEFENSE" 'If the Djinn boosts defense
            AttackType(intTurn) = "DJINNBOOSTDEFENSE"
            AttackDamage(intTurn) = thedamage(1)
            If hoston = True Then
                frmHost.Host.SendData "DJINNDEFENSE" & (intTurn + 2) & vbCrLf
                frmHost.Host.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
            If hoston = False Then
                frmJoin.Client.SendData "DJINNDEFENSE" & (intTurn + 2) & vbCrLf
                frmJoin.Client.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
        Case "ATTACK" 'If the Djinn boosts attack
            AttackType(intTurn) = "DJINNBOOSTAP"
            AttackDamage(intTurn) = thedamage(1)
            If hoston = True Then
                frmHost.Host.SendData "DJINNATTACK" & (intTurn + 2) & vbCrLf
                frmHost.Host.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
            If hoston = False Then
                frmJoin.Client.SendData "DJINNATTACK" & (intTurn + 2) & vbCrLf
                frmJoin.Client.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
        Case "RESIST" 'If the Djinn boosts attack
            AttackType(intTurn) = "DJINNBOOSTRESIST"
            AttackDamage(intTurn) = thedamage(1)
            If hoston = True Then
                frmHost.Host.SendData "DJINNRESIST" & (intTurn + 2) & vbCrLf
                frmHost.Host.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
            If hoston = False Then
                frmJoin.Client.SendData "DJINNRESIST" & (intTurn + 2) & vbCrLf
                frmJoin.Client.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
        Case "DROPRESIST" 'If the Djinn boosts attack
            AttackType(intTurn) = "DJINNREDUCERESIST"
            If hoston = True Then
                frmHost.Host.SendData "DJINNDROPRESIST" & (intTurn + 2) & vbCrLf
                frmHost.Host.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
            If hoston = False Then
                frmJoin.Client.SendData "DJINNDROPRESIST" & (intTurn + 2) & vbCrLf
                frmJoin.Client.SendData "DMG" & (intTurn + 2) & thedamage(1) & vbCrLf
                DoEvents
            End If
        End Select
        
        
        Call HideList
        Call DisableChoose
        
    ElseIf Djinn(intDjinn).State = 1 Then 'If the Djinn is on Standby
    
        AttackType(intTurn) = "DJINNSET"
        AttackDamage(intTurn) = 0
        
        Djinn(intDjinn).State = 0 'Set Djinn
        intDjinnStandby(intTurn) = intDjinnStandby(intTurn) - 1 'Decease # of Djinn on standby
        If intDjinnStandby(intTurn) < 0 Then intDjinnStandby(intTurn) = 0
        
        If hoston = True Then
            frmHost.Host.SendData "SETDJINN" & (intTurn + 2) & vbCrLf 'I'm waiting to set Djinn
            DoEvents
        Else
            frmJoin.Client.SendData "SETDJINN" & (intTurn + 2) & vbCrLf
            DoEvents
        End If
        Call HideList
        
        Call DisableChoose
    Else
        MsgBox "The Djinn is resting."
    End If

End Sub
Public Sub ComSummon()
On Error Resume Next
    AttackType(intTurn) = "SUMMON"
    
    iSummonLevel(intTurn) = CInt(strSumDjinn(lstSummon.ListIndex + 1)) 'Level of the summon
    
    AttackDamage(intTurn) = SummonDamage(lstSummon.ListIndex + 1) 'Get damage that Summon does
    
    Dim ResetDjinn As Integer 'Put Djinn back on Set
    ResetDjinn = 1
    For i = 1 To 10
        If strDjinnName(i) <> "" Then 'If the Djinn exists
            If bDjinnSet(i) = 1 And ResetDjinn <= iSummonLevel(intTurn) Then 'If the Djinn is on standby and there have been less than or equal to djinn reset then the summon's level
                bDjinnSet(i) = 2 'Put Djinn to rest
                ResetDjinn = ResetDjinn + 1
            End If
        End If
    Next 'i
    
    bWaitToResetDjinn = True
    
    intDjinnStandby(intTurn) = intDjinnStandby(intTurn) - ResetDjinn
    
    If intDjinnStandby(intTurn) < 0 Then intDjinnStandby(intTurn) = 0
    
    If hoston = True Then

        frmHost.Host.SendData "SUMMON" & (intTurn + 2) & iSummonLevel(intTurn) & vbCrLf 'I'm waiting to unleash a summon
        frmHost.Host.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        frmHost.Host.SendData "OPREADY" & (intTurn + 2) & vbCrLf

    Else
        frmJoin.Client.SendData "SUMMON" & (intTurn + 2) & iSummonLevel(intTurn) & vbCrLf
        frmJoin.Client.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        frmJoin.Client.SendData "OPREADY" & (intTurn + 2) & vbCrLf
    End If
    
    Call DisableChoose

End Sub
Private Function GetDirName(strPsy As String) As String
On Error Resume Next
If strPsy = "F" Then
    strPsy = "Fire"
ElseIf strPsy = "E" Then
    strPsy = "Earth"
ElseIf strPsy = "W" Then
    strPsy = "Water"
ElseIf strPsy = "N" Then
    strPsy = "Wind"
ElseIf strPsy = "H" Then
    strPsy = "Heart"
Else
    strPsy = "Dark"
End If
GetDirName = strPsy

End Function
Sub WriteCharType(ByVal PowOrResist As String, ByVal Player As Integer, ByVal intChange As Integer)
On Error Resume Next
If Player = 1 Then
    If PowOrResist = "POWER" Then
        Select Case strTypo
        Case "E"
            intEarthPower(1) = intEarthPower(1) + intChange
        Case "F"
            intFirePower(1) = intFirePower(1) + intChange
        Case "N"
            intWindPower(1) = intWindPower(1) + intChange
        Case "W"
            intWaterPower(1) = intWaterPower(1) + intChange
        Case "H"
            intHeartPower(1) = intHeartPower(1) + intChange
        Case "D"
            intDarkPower(1) = intDarkPower(1) + intChange
        End Select
    Else
        Select Case strTypo
        Case "E"
            intEarthResist(1) = intEarthResist(1) + intChange
        Case "F"
            intFireResist(1) = intFireResist(1) + intChange
        Case "N"
            intWindResist(1) = intWindResist(1) + intChange
        Case "W"
            intWaterResist(1) = intWaterResist(1) + intChange
        Case "H"
            intHeartResist(1) = intHeartResist(1) + intChange
        Case "D"
            intDarkResist(1) = intDarkResist(1) + intChange
        End Select
    End If
Else
    If PowOrResist = "POWER" Then
        Select Case strTypo
        Case "E"
            intEarthPower(2) = intEarthPower(2) + intChange
        Case "F"
            intFirePower(2) = intFirePower(2) + intChange
        Case "N"
            intWindPower(2) = intWindPower(2) + intChange
        Case "W"
            intWaterPower(2) = intWaterPower(2) + intChange
        Case "H"
            intHeartPower(2) = intHeartPower(2) + intChange
        Case "D"
            intDarkPower(2) = intDarkPower(2) + intChange
        End Select
    Else
        Select Case strTypo
        Case "E"
            intEarthResist(2) = intEarthResist(2) + intChange
        Case "F"
            intFireResist(2) = intFireResist(2) + intChange
        Case "N"
            intWindResist(2) = intWindResist(2) + intChange
        Case "W"
            intWaterResist(2) = intWaterResist(2) + intChange
        Case "H"
            intHeartResist(2) = intHeartResist(2) + intChange
        Case "D"
            intDarkResist(2) = intDarkResist(2) + intChange
        End Select
    End If
End If
End Sub
Sub SendAttack()
On Error Resume Next

    AttackDamage(intTurn) = Damage
    
    bOReady(intTurn) = True
    
    
    If hoston = True Then
        frmHost.Host.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf 'Damage
        frmHost.Host.SendData "OPREADY" & (intTurn + 2) & vbCrLf 'I'm Ready
        If atType = 0 Then
            frmHost.Host.SendData "DOATTACK" & (intTurn + 2) & vbCrLf 'I'm Going To Attack
        ElseIf atType = 1 Then
            frmHost.Host.SendData "DOCRITICAL" & (intTurn + 2) & vbCrLf 'I'm critical attacking
        Else
            frmHost.Host.SendData "DOSPECIAL" & (intTurn + 2) & vbCrLf 'I'm special attacking
        End If
    Else
        frmJoin.Client.SendData "DMG" & (intTurn + 2) & AttackDamage(intTurn) & vbCrLf
        frmJoin.Client.SendData "OPREADY" & (intTurn + 2) & vbCrLf
        If atType = 0 Then
            frmJoin.Client.SendData "DOATTACK" & (intTurn + 2) & vbCrLf
        ElseIf atType = 1 Then
            frmJoin.Client.SendData "DOCRITICAL" & (intTurn + 2) & vbCrLf
        Else
            frmJoin.Client.SendData "DOSPECIAL" & (intTurn + 2) & vbCrLf
        End If
    End If
    
    Call DisableChoose
End Sub
Sub subSelectTarget()
Call HideList
imgYou(2).MousePointer = 99
imgYou(3).MousePointer = 99
imgYou(6).MousePointer = 99
imgYou(7).MousePointer = 99
SelectTarget = True
lblSelectTarget.Visible = True
End Sub
