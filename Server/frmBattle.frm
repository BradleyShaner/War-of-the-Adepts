VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBattle 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battle Editor"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opBoss 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Boss"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton opRandom 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Random"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame framRandom 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Random Battle Editor"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin MSComDlg.CommonDialog comDiag 
         Left            =   3480
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "*.gif"
         FileName        =   "*.gif"
         Filter          =   "*.gif"
         InitDir         =   "app.path"
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   57
         Top             =   4080
         Width           =   735
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   56
         Top             =   3840
         Width           =   735
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   55
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox txtCoins 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   4560
         MaxLength       =   4
         TabIndex        =   34
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtCoins 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   33
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtCoins 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   4
         TabIndex        =   32
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtPicture 
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   29
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtPicture 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   28
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtPicture 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtAI 
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   26
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtAI 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   25
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtAI 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   3
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   23
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   22
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   3
         TabIndex        =   21
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtAP 
         Height          =   285
         Index           =   3
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   20
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtAP 
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   19
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtAP 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   3
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Index           =   3
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   17
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COINS:"
         Height          =   195
         Index           =   9
         Left            =   2160
         TabIndex        =   31
         Top             =   4200
         Width           =   540
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PICTURE:"
         Height          =   195
         Index           =   8
         Left            =   2160
         TabIndex        =   11
         Top             =   3600
         Width           =   750
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AI (Type one of the following: ATTACK, PSYNERGY, HEAL, DEFEND):"
         Height          =   195
         Index           =   7
         Left            =   315
         TabIndex        =   10
         Top             =   3000
         Width           =   5085
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEFENSE:"
         Height          =   195
         Index           =   6
         Left            =   2160
         TabIndex        =   9
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP:"
         Height          =   195
         Index           =   5
         Left            =   2280
         TabIndex        =   8
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   7
         Top             =   1200
         Width           =   270
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME:"
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   510
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enemy 3:"
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enemy 2:"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enemy 1:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame framBoss 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Boss Editor"
      Height          =   5055
      Left            =   120
      TabIndex        =   30
      Top             =   360
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtNextTile 
         Height          =   285
         Left            =   2520
         TabIndex        =   60
         Text            =   "0"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdBossBrowse 
         Caption         =   "&Browse"
         Height          =   255
         Left            =   1320
         TabIndex        =   58
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox txtNextMap 
         Height          =   285
         Left            =   2520
         TabIndex        =   54
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtTalk2 
         Height          =   285
         Left            =   2520
         TabIndex        =   52
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtCoins 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   360
         MaxLength       =   4
         TabIndex        =   50
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtPicture 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   48
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtAI 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   47
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   4
         Left            =   360
         MaxLength       =   3
         TabIndex        =   46
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtAP 
         Height          =   285
         Index           =   4
         Left            =   360
         MaxLength       =   3
         TabIndex        =   45
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Index           =   4
         Left            =   360
         MaxLength       =   3
         TabIndex        =   44
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   43
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtTalk 
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label lblGen 
         BackStyle       =   0  'Transparent
         Caption         =   "TILE ON NEXT MAP TO LOAD (Input 999 if this is the last map)"
         Height          =   435
         Index           =   20
         Left            =   2520
         TabIndex        =   59
         Top             =   2040
         Width           =   3450
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NEXT MAP TO LOAD:"
         Height          =   195
         Index           =   19
         Left            =   2520
         TabIndex        =   53
         Top             =   1440
         Width           =   1620
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FINISH MESSAGE:"
         Height          =   195
         Index           =   11
         Left            =   2520
         TabIndex        =   51
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COINS:"
         Height          =   195
         Index           =   18
         Left            =   480
         TabIndex        =   49
         Top             =   4440
         Width           =   540
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PICTURE (No filepath)"
         Height          =   195
         Index           =   17
         Left            =   480
         TabIndex        =   42
         Top             =   3840
         Width           =   1605
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AI:"
         Height          =   195
         Index           =   16
         Left            =   480
         TabIndex        =   41
         Top             =   3240
         Width           =   195
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEFENSE:"
         Height          =   195
         Index           =   15
         Left            =   480
         TabIndex        =   40
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP:"
         Height          =   195
         Index           =   14
         Left            =   480
         TabIndex        =   39
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         Height          =   195
         Index           =   13
         Left            =   480
         TabIndex        =   38
         Top             =   1440
         Width           =   270
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME:"
         Height          =   195
         Index           =   12
         Left            =   480
         TabIndex        =   37
         Top             =   840
         Width           =   510
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dialogue (Use \n to detonate a new line)"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2880
      End
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBossBrowse_Click()
On Error Resume Next
comDiag.ShowOpen
txtPicture(Index + 1).Text = comDiag.FileName
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
On Error Resume Next
comDiag.ShowOpen
txtPicture(Index + 1).Text = comDiag.FileName
End Sub

Private Sub opBoss_Click()
framRandom.Visible = False
framBoss.Visible = True
End Sub

Private Sub opRandom_Click()
framRandom.Visible = True
framBoss.Visible = False
End Sub

Private Sub txtAP_Change(Index As Integer)
On Error Resume Next
If Admin = False Then
    Dim intCoins As Integer
    intCoins = (0.3 * CInt(txtHP(Index).Text)) + (0.4 * CInt(txtAP(Index).Text)) + (0.4 * CInt(txtDefense(Index).Text))
    txtCoins(Index).Text = CStr(intCoins)
End If
End Sub

Private Sub txtDefense_Change(Index As Integer)
On Error Resume Next
If Admin = False Then
    Dim intCoins As Integer
    intCoins = (0.3 * CInt(txtHP(Index).Text)) + (0.4 * CInt(txtAP(Index).Text)) + (0.4 * CInt(txtDefense(Index).Text))
    txtCoins(Index).Text = CStr(intCoins)
End If
End Sub

Private Sub txtHP_Change(Index As Integer)
On Error Resume Next
If Admin = False Then
    Dim intCoins As Integer
    intCoins = (0.3 * CInt(txtHP(Index).Text)) + (0.4 * CInt(txtAP(Index).Text)) + (0.4 * CInt(txtDefense(Index).Text))
    txtCoins(Index).Text = CStr(intCoins)
End If
End Sub

