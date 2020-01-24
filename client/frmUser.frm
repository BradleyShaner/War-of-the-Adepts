VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Configuration Tool"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   628
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtCaptionPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   52
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdBuyDjinn 
      Caption         =   "Buy Djinn"
      Height          =   255
      Left            =   7320
      TabIndex        =   50
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer timeTimeout 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5880
      Top             =   4560
   End
   Begin VB.ListBox lstPsy 
      Height          =   2010
      ItemData        =   "frmUser.frx":0000
      Left            =   6240
      List            =   "frmUser.frx":0002
      TabIndex        =   45
      Top             =   240
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set Djinn"
      Height          =   255
      Left            =   6240
      TabIndex        =   43
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstDjinn 
      Height          =   2010
      ItemData        =   "frmUser.frx":0004
      Left            =   6240
      List            =   "frmUser.frx":0006
      TabIndex        =   42
      Top             =   240
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton cmdBuyItem 
      Caption         =   "Buy Item"
      Height          =   255
      Left            =   6240
      TabIndex        =   41
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstItem 
      Height          =   2010
      ItemData        =   "frmUser.frx":0008
      Left            =   6240
      List            =   "frmUser.frx":000A
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   195
      Left            =   8160
      TabIndex        =   38
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer timeshrink 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   8280
      Top             =   3120
   End
   Begin VB.CommandButton cmdPsynergy 
      Caption         =   "View Psynergy"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6360
      Picture         =   "frmUser.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdDjinn 
      Caption         =   "Djinn Config"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6360
      Picture         =   "frmUser.frx":2BCA
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "Item Shop"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmUser.frx":5365
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox cmbCharPic 
      Height          =   315
      ItemData        =   "frmUser.frx":681F
      Left            =   1920
      List            =   "frmUser.frx":6850
      TabIndex        =   14
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdVal 
      Caption         =   "Log In"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtusername 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      MaxLength       =   12
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtpassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtnewpassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtnewusername 
      Height          =   285
      Left            =   120
      MaxLength       =   12
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create New User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock User 
      Left            =   4800
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9888
   End
   Begin VB.Label lblgen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   34
      Left            =   1920
      TabIndex        =   51
      Top             =   4080
      Width           =   1155
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   33
      Left            =   3360
      TabIndex        =   49
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   32
      Left            =   960
      TabIndex        =   48
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Character:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   31
      Left            =   2280
      TabIndex        =   47
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Weapon:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   30
      Left            =   0
      TabIndex        =   46
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Djinn on Standby!"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   29
      Left            =   6240
      TabIndex        =   44
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "No Item Selected!"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   28
      Left            =   6240
      TabIndex        =   40
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit User:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   27
      Left            =   6240
      TabIndex        =   34
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   26
      Left            =   3960
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   25
      Left            =   2520
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   24
      Left            =   720
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Level:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   3240
      TabIndex        =   30
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Djinn:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   1800
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Coins:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   0
      TabIndex        =   28
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "None Selected"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   4440
      TabIndex        =   27
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Class:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   4440
      TabIndex        =   26
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Login: (Note: You can only use lowercase letters)"
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
      Height          =   255
      Index           =   18
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "No character selected!"
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
      Height          =   975
      Index           =   17
      Left            =   4080
      TabIndex        =   24
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Description:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   23
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   16
      Left            =   5520
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   15
      Left            =   3720
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   2280
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Disconnects:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   4320
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Losses:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   2880
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Wins:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   1560
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   840
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Rating:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Character:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Not logged in."
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   6
      Left            =   3360
      TabIndex        =   12
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My User Name:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "My Password:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Create A New User:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strdata() As String
Dim curOne As Byte



Dim sCurItem As String
Dim iCurItem As Integer
Dim lostEMAIL As String
Dim lostPW As String

Private Sub cmbCharPic_Click()
Debug.Print cmbCharPic.Text
If cmbCharPic.Text = "Isaac" Then
lblgen(17).Caption = "Isaac: An all around character with good Attack."
lblgen(20).Caption = "Earth"
End If
If cmbCharPic.Text = "Garret" Then
lblgen(17).Caption = "Garret: Good HP, bad PP, average Attack."
lblgen(20).Caption = "Fire"
End If
If cmbCharPic.Text = "Jenna" Then
lblgen(17).Caption = "Jenna: Good PP, Average HP, Average Attack, Psynergy suffers 1/5 damage loss."
lblgen(20).Caption = "Fire"
End If
If cmbCharPic.Text = "Ivan" Then
lblgen(17).Caption = "Ivan: Good PP, Bad HP, Bad Attack, Gains 1/4 Psynergy Damage Bonus."
lblgen(20).Caption = "Wind"
End If
If cmbCharPic.Text = "Mia" Then
lblgen(17).Caption = "Mia: Great PP, Bad Attack, Average HP, Suffers 1/4 Psynergy Damage Loss."
lblgen(20).Caption = "Water"
End If
If cmbCharPic.Text = "Sheba" Then
lblgen(17).Caption = "Sheba: Average HP, Average PP, Bad Attack. Gains 1/3 Psynergy Damage Bonus."
lblgen(20).Caption = "Wind"
End If
If cmbCharPic.Text = "Felix" Then
lblgen(17).Caption = "Felix: Great Attack, Average HP, Bad PP."
lblgen(20).Caption = "Heart"
End If
If cmbCharPic.Text = "Alex" Then
lblgen(17).Caption = "Alex: Good HP, Good Attack, Bad PP.  Gets Evasion Bonus."
lblgen(20).Caption = "Water"
End If
If cmbCharPic.Text = "Saturos" Then
lblgen(17).Caption = "Saturos: Great HP, Great Attack, Bad PP.  Psynergy Cost Doubled."
lblgen(20).Caption = "Fire"
End If
If cmbCharPic.Text = "Menardi" Then
lblgen(17).Caption = "Menardi: Good HP, Average PP, Average Attack."
lblgen(20).Caption = "Fire"
End If
If cmbCharPic.Text = "Kraden" Then
lblgen(17).Caption = "Kraden: Exceptional PP, Bad HP, Bad Attack.  Gains 1/3 Psynergy Damage Bonus, Chance of 'Alchemy Bonus'"
lblgen(20).Caption = "Dark"
End If
If cmbCharPic.Text = "Babi" Then
lblgen(17).Caption = "Babi: Secret character.  Requires 5 wins to unlock!"
lblgen(20).Caption = "?????"
End If
If cmbCharPic.Text = "Caption Contest Character" Then
txtCaptionPW.Visible = True
lblgen(34).Visible = True
lblgen(17).Caption = "This is for winners of the GSA Caption Contest only.  Please enter the password in your award picture."
lblgen(20).Caption = "?????"
Else
If cmbCharPic.Text = "Guard" Then
lblgen(17).Caption = "Guard: No Psynergy, Weak Attack, Weak Defense.  EXCEPTIONAL leveling-up stats."
lblgen(20).Caption = "Normal"
End If
If cmbCharPic.Text = "Gladiator" Then
lblgen(17).Caption = "Gladiator: Self proclaimed dominator of low-level play.  Great initial Attack and Defense.  No Psynergy."
lblgen(20).Caption = "Normal"
End If
txtCaptionPW.Visible = False
lblgen(34).Visible = False
End If
End Sub

Private Sub cmdBuyItem_Click()
On Error Resume Next
Dim iCoins As Integer
Dim iValue As Integer
iCoins = CInt(strCoins)
iValue = CInt(strItemCoins(lstItem.ListIndex + 1))
If iCoins >= iValue Then
lblgen(6).Caption = "Item bought!"
iCoins = iCoins - iValue
lblgen(24).Caption = iCoins
Timeout (2)
lblgen(6).Caption = "Updating Stats in server"
User.SendData "NEWITEMUSER" & strMyUserName & vbCrLf
User.SendData "NEWITEMCOINS" & icons & vbCrLf
User.SendData "NEWITEMNAME" & lstItem.Text & vbCrLf
DoEvents
Else
MsgBox "Sorry, you do not have enough funds to buy this item."
End If
End Sub

Private Sub cmdChat_Click()
'If chatLoaded = False Then
'lblgen(6).Caption = "Rejoining chat, please wait..." '
'User.Connect IKILLKENNYIP, 9888
'frmChat.Show 1
'Else
'lblgen(6).Caption = "Chat window already open!"
'End If
If AmIKilled = False Then
frmChat.Show
Else
MsgBox "You are not allowd back in the chat!"
End If
End Sub

Private Sub cmdDjinn_Click()
timeshrink.Enabled = True
curOne = 1
End Sub

Private Sub cmdExit_Click()
cmdExit.Visible = False
lstPsy.Visible = False
lblgen(29).Visible = False
lstDjinn.Visible = False
cmdSet.Visible = False
cmdItem.Visible = True
cmdBuyItem.Visible = False
lblgen(28).Visible = False
lstItem.Visible = False
cmdPsynergy.Visible = True
cmdDjinn.Visible = True
cmdItem.Height = 97
cmdItem.Width = 97
cmdDjinn.Height = 97
cmdDjinn.Width = 97
cmdPsynergy.Height = 97
cmdPsynergy.Width = 97
frmUser.Width = 9090
End Sub

Private Sub cmdItem_Click()
timeshrink.Enabled = True
curOne = 0
End Sub

Private Sub cmdLostPW_Click()
On Error Resume Next
lostEMAIL = InputBox("What is your e-mail address?")
lostPW = InputBox("What is your username?")
LostPassword = True
NewUser = False
If frmHost.Host.State <> sckClosed Then
User.Close
End If
User.Connect IKILLKENNYIP, 9888
End Sub

Private Sub cmdPsynergy_Click()
timeshrink.Enabled = True
curOne = 2
End Sub

Private Sub cmdSet_Click()
On Error Resume Next
If iDjinnSet(lstDjinn.ListIndex + 1) = 0 Then
iDjinnSet(lstDjinn.ListIndex + 1) = 1
lblgen(29).Caption = "Djinn Set!"
Else
iDjinnSet(lstDjinn.ListIndex + 1) = 0
lblgen(29).Caption = "Djinn on Standby!"
End If
End Sub

Private Sub cmdVal_Click()
On Error Resume Next
strMyUserName = txtUserName.Text
strMyPassWord = txtPassword.Text
If User.State <> sckClosed Then
User.Close
Timeout (2)
End If
lblgen(6).Caption = "Attempting to log in"
NewUser = False
LostPassword = False
User.Connect IKILLKENNYIP, 9888
timeTimeout.Enabled = True
End Sub

Private Sub Command1_Click()
'On Error Resume Next
LostPassword = False
NewUser = True
If frmHost.Host.State <> sckClosed Then
User.Close
DoEvents
End If
If cmbCharPic.Text <> "" And cmbCharPic.Text <> "Babi" Then
With cmbCharPic
    If .Text = "Guard" Or .Text = "Gladiator" Or .Text = "Isaac" Or .Text = "Garret" Or .Text = "Ivan" Or .Text = "Mia" Or .Text = "Jenna" Or .Text = "Saturos" Or .Text = "Menardi" Or .Text = "Alex" Or .Text = "Felix" Or .Text = "Sheba" Then
        User.Connect IKILLKENNYIP, 9888
    Else
        If .Text = "Caption Contest Character" And txtCaptionPW.Text = "alpha" Then
            User.Connect IKILLKENNYIP, 9888
        Else
            lblgen(6).Caption = "Please reselect a character, and don't modify the text."
        End If
    End If
End With
Else
lblgen(6).Caption = "There was an error in creating your new account.  Please select your class and picture."
End If
End Sub

Private Sub Form_Load()

'If LoggedIn = True Then
'cmdDjinn.Enabled = True
'cmdPsynergy.Enabled = True
'cmdItem.Enabled = True
'cmdChat.Enabled = True
'    lblgen(9).Visible = True
'    lblgen(10).Visible = True
'    lblgen(11).Visible = True
'    lblgen(12).Visible = True
'    lblgen(13).Visible = True
'    lblgen(26).Visible = True
'    lblgen(25).Visible = True
'    lblgen(24).Visible = True
'    lblgen(23).Visible = True
'    lblgen(22).Visible = True
'    lblgen(21).Visible = True
'    lblgen(16).Visible = True
'    lblgen(15).Visible = True
'    lblgen(14).Visible = True
'    lblgen(30).Visible = True
'    lblgen(31).Visible = True
'    lblgen(32).Visible = True
'    lblgen(33).Visible = True
'    cmdVal.Enabled = False
'    txtusername.Enabled = False
'    txtpassword.Enabled = False
'    cmdLostPW.Enabled = False
'    Command1.Enabled = False
'End If

frmUser.Picture = frmIntro.Picture

If User.LocalIP = "192.168.0.2" Then
    IKILLKENNYIP = User.LocalIP
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Cancel = 1
End Sub

Private Sub imgIsaac_Click()
Beep
lblEasterEgg.Visible = True
End Sub

Private Sub lstDjinn_Click()
lblgen(28).Caption = strDjinnDesc(lstDjinn.ListIndex + 1)
If iDjinnSet(lstDjinn.ListIndex + 1) = 0 Then
lblgen(29).Caption = "Djinn on Standby!"
Else
lblgen(29).Caption = "Djinn Set!"
End If
End Sub

Private Sub lstItem_Click()
lblgen(28).Caption = strItemDesc(lstItem.ListIndex + 1)
End Sub

Private Sub timeshrink_Timer()
frmUser.Width = frmUser.Width + 10
If cmdItem.Height >= 14 Then
cmdItem.Height = cmdItem.Height - 3
cmdPsynergy.Height = cmdPsynergy.Height - 3
cmdDjinn.Height = cmdDjinn.Height - 3
Else
    If cmdItem.Width >= 6 Then
    cmdItem.Width = cmdItem.Width - 3
    cmdPsynergy.Width = cmdPsynergy.Width - 3
    cmdDjinn.Width = cmdDjinn.Width - 3
    Else
    cmdItem.Visible = False
    cmdPsynergy.Visible = False
    cmdDjinn.Visible = False
    cmdExit.Visible = True
    timeshrink.Enabled = False
    If curOne = 0 Then
    lblgen(27).Caption = "Item Shop:"
    lblgen(28).Visible = True
    lstItem.Visible = True
    cmdBuyItem.Visible = True
    End If
    If curOne = 1 Then
    lblgen(27).Caption = "Djinn Configuration:"
    lstDjinn.Visible = True
    cmdSet.Visible = True
    lblgen(28).Visible = True
    lblgen(29).Visible = True
    End If
    If curOne = 2 Then
    lblgen(27).Caption = "View Psynergy:"
    lstPsy.Visible = True
    End If
    End If
End If
End Sub

Private Sub timeTimeout_Timer()
lblgen(6).Caption = "Connection timed out.  Please consult the Status of the server by clicking on the Status button on the main window."
timeTimeout.Enabled = False
User.Close
End Sub

Private Sub txtCaptionPW_Change()
txtCaptionPW.Text = LCase(txtCaptionPW.Text)
txtCaptionPW.SelStart = Len(txtCaptionPW.Text)
If txtCaptionPW.Text = "alpha" Then
    lblgen(17).Caption = "Lizard Man is a fearsome beast with unheard of strength."
    lblgen(20).Caption = "Water"
End If
End Sub

Private Sub txtnewpassword_Change()
txtnewpassword.Text = LCase(txtnewpassword.Text)
txtnewpassword.SelStart = Len(txtnewpassword.Text)
End Sub

Private Sub txtnewusername_Change()
txtnewusername.Text = LCase(txtnewusername.Text)
txtnewusername.SelStart = Len(txtnewusername.Text)
End Sub

Private Sub txtpassword_Change()
txtPassword.Text = LCase(txtPassword.Text)
txtPassword.SelStart = Len(txtPassword.Text)
End Sub

Private Sub txtUserName_Change()
txtUserName.Text = LCase(txtUserName.Text)
txtUserName.SelStart = Len(txtUserName.Text)
End Sub

Private Sub User_Connect()
On Error Resume Next
Beep
timeTimeout.Enabled = False
'timeTimeout.Enabled = True
lblgen(6).Caption = "Logged into server"
If disconnect = False Then
If txtUserName.Text = "root" Then
User.SendData "USER" & "mike" & vbCrLf & "PASS" & "qwerty" & vbCrLf
Else
    If LoggedIn = False Then
        If NewUser = False And LostPassword = False And WinBattle = False Then
        Call sendpass
        Call SendError
        End If
        If NewUser = True Then
        Call createuser
        End If
        If LostPassword = True Then
        Call requestpw
        End If
    Else
        If WinBattle = True Then
        Call StatSend
        Else
        Call Populate
        End If
    End If
End If
Else
User.SendData "DISC" & strOpponent & vbCrLf
End If
End Sub

Private Sub User_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strdatao As String
User.GetData strdatao
strdata = Split(strdatao, vbCrLf, -1, vbTextCompare)
For i = 0 To UBound(strdata)
If Left$(strdata(i), 4) = "GOOD" Then
timeTimeout.Enabled = False
lblgen(6).Caption = "Log-in successful!"
    'cmdVal.Enabled = False
    frmHost.cmdListen.Enabled = True
    frmJoin.cmdListen.Enabled = True
    txtUserName.Enabled = False
    txtPassword.Enabled = False
    cmdVal.Enabled = False
End If
If Left$(strdata(i), 3) = "BAD" Then
timeTimeout.Enabled = False
lblgen(6).Caption = "Bad user name, password, or version of the game!"
    'MsgBox "Bad User Name/Password!"
    Timeout (2)
    'Host.SendData "CLOSE"
    User.Close
End If
If Left$(strdata(i), 4) = "BUSR" Then
timeTimeout.Enabled = False
lblgen(6).Caption = "Could not create a new user.  Try a different user name, and make sure you have something entered in the password text box."
    'MsgBox "Could not create a new user!"
    Timeout (2)
    'Host.SendData "CLOSE"
    User.Close
    lblgen(6).Caption = "Logged out of server."
End If
If Left$(strdata(i), 4) = "GUSR" Then
timeTimeout.Enabled = False
lblgen(6).Caption = "Succesfully created a new user!"
    'MsgBox "New user created!"
    Timeout (2)
    'User.SendData "CLOSE"
    User.Close
    lblgen(6).Caption = "Logged out of server."
End If
If Left$(strdata(i), 3) = "ERR" Then
timeTimeout.Enabled = False
    lblgen(6).Caption = "An error occured, attempting to try again!"
    If NewUser = False And LostPassword = False Then
    Call sendpass
    End If
    If NewUser = True Then
    Call createuser
    End If
    If LostPassword = True Then
    Call requestpw
    End If
End If
If Left$(strdata(i), 4) = "STAT" Then
timeTimeout.Enabled = False
    User.Close
    lblgen(6).Caption = "Stats sent successfully.  Please log-in again to update your stats."
    WinBattle = False
    cmdVal.Enabled = True
    Command1.Enabled = True
    cmdLostPW.Enabled = True
End If

If Left$(strdata(i), 7) = "CURITEM" Then
    sCurItem = Mid(strdata(i), 8, Len(strdata(i)))
    iCurItem = CInt(sCurItem)
End If
If Left$(strdata(i), 8) = "ITEMNAME" Then
    strItemName(iCurItem) = Mid(strdata(i), 9, Len(strdata(i)))
End If
If Left$(strdata(i), 8) = "ITEMDESC" Then
    strItemDesc(iCurItem) = Mid(strdata(i), 9, Len(strdata(i)))
End If
If Left$(strdata(i), 7) = "ITEMDMG" Then
    strItemDamage(iCurItem) = Mid(strdata(i), 8, Len(strdata(i)))
End If
If Left$(strdata(i), 10) = "ITEMSPCDMG" Then
    strItemSpcDamage(iCurItem) = Mid(strdata(i), 11, Len(strdata(i)))
End If
If Left$(strdata(i), 11) = "ITEMSPCDESC" Then
    strItemSpcDesc(iCurItem) = Mid(strdata(i), 12, Len(strdata(i)))
End If
If Left$(strdata(i), 11) = "ITEMSPCTYPE" Then
    strItemSpcDesc(iCurItem) = Mid(strdata(i), 12, Len(strdata(i)))
End If
If Left$(strdata(i), 8) = "ITEMTYPE" Then
    strItemSpcDesc(iCurItem) = Mid(strdata(i), 9, Len(strdata(i)))
End If
If Left$(strdata(i), 9) = "ITEMCOINS" Then
    strItemCoins(iCurItem) = Mid(strdata(i), 10, Len(strdata(i)))
    lstItem.AddItem strItemName(iCurItem) & " Costs: " & strItemCoins(iCurItem) & " coins."
End If

If Left$(strdata(i), 8) = "CURDJINN" Then
    sCurDjinn = Mid(strdata(i), 9, Len(strdata(i)))
    iCurDjinn = CInt(sCurDjinn)
End If
If Left$(strdata(i), 9) = "DJINNNAME" Then
    strDjinnName(iCurDjinn) = Mid(strdata(i), 10, Len(strdata(i)))
    lstDjinn.AddItem strDjinnName(iCurDjinn)
End If
If Left$(strdata(i), 9) = "DJINNDESC" Then
    strDjinnDesc(iCurDjinn) = Mid(strdata(i), 10, Len(strdata(i)))
End If
If Left$(strdata(i), 9) = "DJINNTYPE" Then
    strDjinnType(iCurDjinn) = Mid(strdata(i), 10, Len(strdata(i)))
End If
If Left$(strdata(i), 8) = "DJINNDMG" Then
    strDjinnDamage(iCurDjinn) = Mid(strdata(i), 9, Len(strdata(i)))
End If

If Left$(strdata(i), 6) = "CURPSY" Then
    sCurPsy = Mid(strdata(i), 7, Len(strdata(i)))
    iCurPsy = CInt(sCurPsy)
End If
If Left$(strdata(i), 7) = "PSYNAME" Then
    strPsyName(iCurPsy) = Mid(strdata(i), 8, Len(strdata(i)))
End If
If Left$(strdata(i), 6) = "PSYDMG" Then
    strPsyDamage(iCurPsy) = Mid(strdata(i), 7, Len(strdata(i)))
End If
If Left$(strdata(i), 7) = "PSYTYPE" Then
    strPsyType(iCurPsy) = Mid(strdata(i), 8, Len(strdata(i)))
End If
If Left$(strdata(i), 5) = "PSYPP" Then
    strPsyPP(iCurPsy) = Mid(strdata(i), 6, Len(strdata(i)))
End If
If Left$(strdata(i), 8) = "PSYDJINN" Then
    strPsyDjinn(iCurPsy) = Mid(strdata(i), 9, Len(strdata(i)))
    lstPsy.AddItem (strPsyName(iCurPsy) & " Djinn Required:" & strPsyDjinn(iCurPsy))
End If

If Left$(strdata(i), 6) = "CURSUM" Then
    sCurSum = Mid(strdata(i), 7, Len(strdata(i)))
    iCurSum = CInt(sCurSum)
End If
If Left$(strdata(i), 7) = "SUMNAME" Then
    strSumName(iCurSum) = Mid(strdata(i), 8, Len(strdata(i)))
End If
If Left$(strdata(i), 8) = "SUMDJINN" Then
    strSumDjinn(iCurSum) = Mid(strdata(i), 9, Len(strdata(i)))
'    frmBattle.lstDjinn.AddItem (strSumName(iCurSum) & " " & strSumDjinn(iCurSum) & " Djinn Needed.")
End If
If Left$(strdata(i), 7) = "SUMDESC" Then
    strSumDesc(iCurSum) = Mid(strdata(i), 8, Len(strdata(i)))
End If

'   For r = 1 To itotal
'        Server.SendData "CURSUM" & r & vbCrLf
'        stotal = GetFromIni(curType & r, "NAME", nsave)
'        Server.SendData "SUMNAME" & stotal & vbCrLf
'        stotal = GetFromIni(curType & r, "DJINN", nsave)
'        Server.SendData "SUMDJINN" & stotal & vbCrLf
'        stotal = GetFromIni(curType & r, "DESC", nsave)
'        Server.SendData "SUMDESC" & stotal & vbCrLf
'    Next 'r

If Left$(strdata(i), 6) = "RATING" Then
    lblgen(10).Caption = Mid(strdata(i), 7, Len(strdata(i)))
    strRating = Mid(strdata(i), 7, Len(strdata(i)))
End If
If Left$(strdata(i), 4) = "WINS" Then
    lblgen(14).Caption = Mid(strdata(i), 5, Len(strdata(i)))
    strWins = Mid(strdata(i), 5, Len(strdata(i)))
End If
If Left$(strdata(i), 4) = "LOSS" Then
    lblgen(15).Caption = Mid(strdata(i), 5, Len(strdata(i)))
    strLoss = Mid(strdata(i), 5, Len(strdata(i)))
End If
If Left$(strdata(i), 4) = "DISC" Then
    lblgen(16).Caption = Mid(strdata(i), 5, Len(strdata(i)))
    strDisc = Mid(strdata(i), 5, Len(strdata(i)))
End If
If Left$(strdata(i), 5) = "COINS" Then
    lblgen(24).Caption = Mid(strdata(i), 6, Len(strdata(i)))
    strCoins = Mid(strdata(i), 6, Len(strdata(i)))
End If
If Left$(strdata(i), 5) = "DJINN" Then
    lblgen(25).Caption = Mid(strdata(i), 6, Len(strdata(i)))
    strDjinn = Mid(strdata(i), 6, Len(strdata(i)))
End If
If Left$(strdata(i), 4) = "CHAR" Then
    strChar = Mid(strdata(i), 5, Len(strdata(i)))
    lblgen(33).Caption = strChar
End If
If Left$(strdata(i), 3) = "WPN" Then
    strWeapon = Mid(strdata(i), 4, Len(strdata(i)))
    intWeapon = CInt(strWeapon)
    lblgen(32).Caption = strItemName(intWeapon)
End If

If Left$(strdata(i), 3) = "LVL" Then
timeTimeout.Enabled = False
    lblgen(26).Caption = Mid(strdata(i), 4, Len(strdata(i)))
    strLvl = Mid(strdata(i), 4, Len(strdata(i)))
    lblgen(9).Visible = True
    lblgen(10).Visible = True
    lblgen(11).Visible = True
    lblgen(12).Visible = True
    lblgen(13).Visible = True
    lblgen(26).Visible = True
    lblgen(25).Visible = True
    lblgen(24).Visible = True
    lblgen(23).Visible = True
    lblgen(22).Visible = True
    lblgen(21).Visible = True
    lblgen(16).Visible = True
    lblgen(15).Visible = True
    lblgen(14).Visible = True
    lblgen(30).Visible = True
    lblgen(31).Visible = True
    lblgen(32).Visible = True
    lblgen(33).Visible = True
    
'    txtusername.Visible = False
'    txtpassword.Visible = False
'    lblgen(0).Visible = False
'    lblgen(1).Visible = False
    cmdPsynergy.Enabled = True
    cmdItem.Enabled = True
    cmdDjinn.Enabled = True
    cmdChat.Enabled = True
    LoggedIn = True
'    Timeout (2)
    'Host.SendData "CLOSE"
End If

    If Left$(strdata(i), 10) = "ITEMCONFIRM" Then
        lblgen(6).Caption = "Item updated!"
    End If

Next 'i
Exit Sub
err:
lblgen(6).Caption = "There was an error communicating with the main server."
'MsgBox "There was an error communicating with Ikillkenny's Main Server.  Please try again."
End Sub

Sub sendpass()
On Error Resume Next
'Timeout (2)
User.SendData "USER" & txtUserName.Text & vbCrLf
    lblgen(6).Caption = "Sending user name"
User.SendData "VERS" & Version & vbCrLf
    lblgen(6).Caption = "Sending version."
    'DoEvents
'Timeout (4)
    User.SendData "PASS" & txtPassword.Text & vbCrLf
    'DoEvents
lblgen(6).Caption = "Sending password"
End Sub
Sub createuser()
On Error Resume Next
User.SendData "NEWUSER" & frmUser.txtnewusername.Text & vbCrLf
DoEvents
User.SendData "CHAR" & cmbCharPic.Text & vbCrLf
DoEvents
User.SendData "NEWPW" & frmUser.txtnewpassword.Text & vbCrLf
DoEvents
NewUser = False
End Sub
Sub requestpw()
On Error Resume Next
DoEvents
User.SendData "LOSTEMAIL" & lostEMAIL & vbCrLf
DoEvents
'Timeout (2)
DoEvents
User.SendData "LOSTUSER" & lostPW & vbCrLf
DoEvents
'Timeout (2)
User.Close
lblgen(6).Caption = "Lost Password request sent."
LostPassword = False
End Sub
Sub StatSend()

End Sub
Sub Populate()
On Error Resume Next
User.SendData "CHATPOPN" & strMyUserName & vbCrLf
User.SendData "CHATPOPR" & strRating & vbCrLf
End Sub
Sub SendError()
'On Error Resume Next
Dim nsave As String
nsave = App.Path & "\userdata.ini"
strerror = GetFromIni("ERROR", "ERROR", nsave)
strsource = GetFromIni("ERROR", "SOURCE", nsave)
strNum = GetFromIni("ERROR", "NUMBER", nsave)
If strerror <> "" Then
    User.SendData "ERRORDESC" & strerror & vbCrLf
    User.SendData "ERRORSOURCE" & strsource & vbCrLf
    User.SendData "ERRORNUM" & strNum & vbCrLf
End If
Call WriteIni("ERROR", "ERROR", "", nsave)
Call WriteIni("ERROR", "SOURCE", "", nsave)
Call WriteIni("ERROR", "NUMBER", "", nsave)


End Sub
