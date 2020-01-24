VERSION 5.00
Begin VB.Form frmQuest 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Overall Quest Editor"
   ClientHeight    =   2295
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEnd 
      Height          =   285
      Left            =   120
      MaxLength       =   99
      TabIndex        =   9
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox txtCoins 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "200"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   120
      MaxLength       =   99
      TabIndex        =   5
      Top             =   1200
      Width           =   5295
   End
   Begin VB.TextBox txtTile 
      Height          =   285
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   120
      MaxLength       =   99
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Message:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reward (In Coins) For Winning:"
      Height          =   195
      Index           =   3
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   2205
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level Description:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tile To Place Character:"
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Level To Load:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1410
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuClose_Click()
Me.Hide
End Sub

Private Sub mnuOpen_Click()
Dim strSave As String
strSave = InputBox("Enter the name of the file to save without an extension.  Ex.: Quest_For_The_Sol_Blade")
strSave = App.Path & "\" & strSave & ".quest"
txtLvl.Text = UltraDecode("FIRSTLVL", "FIRSTLVLL", strSave)
txtTile.Text = UltraDecode("TILEPOS", "TILEPOSL", strSave)
txtDesc.Text = UltraDecode("DESC", "DESCL", strSave)
txtCoins.Text = UltraDecode("COINS", "COINSL", strSave)
txtEnd.Text = UltraDecode("END", "ENDL", strSave)
End Sub

Private Sub mnuSave_Click()
Dim strSave As String
strSave = InputBox("Enter the name of the file to save without an extension.  Ex.: Quest_For_The_Sol_Blade")
strSave = App.Path & "\" & strSave & ".quest"
Call Encode(txtLvl.Text, "FIRSTLVL", "FIRSTLVLL", strSave)
Call Encode(txtTile.Text, "TILEPOS", "TILEPOSL", strSave)
Call Encode(txtDesc.Text, "DESC", "DESCL", strSave)
Call Encode(txtCoins.Text, "COINS", "COINSL", strSave)
Call Encode(txtEnd.Text, "END", "ENDL", strSave)
End Sub
