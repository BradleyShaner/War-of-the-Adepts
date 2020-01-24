VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUser2 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts - Log In"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   Icon            =   "frmUser2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmUser2.frx":08CA
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timeDownloadLag 
      Enabled         =   0   'False
      Interval        =   25000
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
      Pattern         =   "*.gif"
      TabIndex        =   33
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkSaveLogin 
      BackColor       =   &H00404080&
      Caption         =   "Save User Name/Password"
      Height          =   255
      Left            =   480
      TabIndex        =   32
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton cmdTips 
      Caption         =   "More &Tips"
      Height          =   255
      Left            =   3360
      TabIndex        =   31
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
      ItemData        =   "frmUser2.frx":1194
      Left            =   480
      List            =   "frmUser2.frx":11C5
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   2640
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock User 
      Left            =   2880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9888
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana Ref"
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
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Verdana Ref"
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
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtTOS 
      Height          =   3015
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmUser2.frx":1239
      Top             =   480
      Width           =   2655
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
      TabIndex        =   30
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
      TabIndex        =   29
      Top             =   4080
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
      TabIndex        =   28
      Top             =   360
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   3360
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
      TabIndex        =   14
      Top             =   3000
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   4
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmUser2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iCurCust As Long
Dim RealChar As Boolean
Dim bWait As Boolean
Dim FileNameList() As String 'array of file names
Dim intCurFile As Integer 'what is the current
Dim b64 As New base64 'initiate base64 class



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
ElseIf cmbCharPic.Text = "Picard" Then
    lblChar(1).Caption = "Picard: A powerful sea pirate that specializes in his high power and resist."
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
ElseIf cmbCharPic.Text = "Purple Picard" Then
    lblChar(1).Caption = "Purple Picard: Picard, but with gnarly purple hair."
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
ElseIf intCurCust <> 999 Then
    lblChar(1).Caption = CustomChar(intCurCust).Name & ": " & CustomChar(intCurCust).Description
    Dim strCurCustType As String
    strCurCustType = GetFullElementalType(CustomChar(intCurCust).Type)
    lblChar(2).Caption = strCurCustType
Else
    RealChar = False 'not a real character
End If

End Sub

Private Sub cmdLogin_Click()
On Error GoTo err

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
strMyUserName = txtUserName.Text
strMyPassWord = txtPassword.Text

If bNewChar = False Then

lblStatus(1).Caption = "Logging in... (if not connected within 15 seconds, hit the Log In button again)"

'Comment out below for LADDER TOURNAMENT
User.SendData "USER" & txtUserName.Text & vbCrLf
User.SendData "VERS" & Version & vbCrLf
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
        Dim intCurCustNum As Long
        intCurCustNum = FindWhichCharacter(cmbCharPic.Text)
        
        
        lblStatus(1).Caption = "Creating character..."
         
        User.SendData "NEWUSER" & txtUserName.Text & vbCrLf
        
        If intCurCustNum = 999 Then
            User.SendData "CHAR" & cmbCharPic.Text & vbCrLf
        Else
            User.SendData "CUSTCHAR" & intCurCustNum & vbCrLf
        End If
        
        User.SendData "NEWPW" & txtPassword.Text & vbCrLf
         
    Else
        If bSpoof = False Then
        MsgBox "Error: The character you have selected does not exist."
        End If
    End If

End If

Exit Sub
err:
MsgBox "Error: Not connected to the server.  If this problem continues, go to the Options Menu and choose Server Status to see if the server is up."

End Sub

Private Sub cmdProceed_Click()
Unload Me
frmChat.Show
End Sub

Private Sub cmdTips_Click()
'Cycles through random tips of the day
Dim rndInt As Integer
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
End Sub

Private Sub FileTransfer_Connect()
On Error Resume Next
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

Call WriteIni("debug", t, test, App.Path & "\debug.ini")

If Left$(test, 3) = "LST" Then

    FileNameList = Split(Mid$(test, 4, Len(test)), ",", -1, vbTextCompare)
    intCurFile = 0
    Call DownloadNextFile

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
    If User.State = sckClosed Then
        User.Connect IKILLKENNYIP, 9888
        lblStatus(1).Caption = "Attempting to connect to the server."
    End If

'Dim rndInt As Integer
'rndInt = Int(Rnd * 5)
'Select Case rndInt
'Case 1
    lblTip(1).Caption = "Press 'R' in the Online Town to reset your character."
'Case 2
'    lblTip(1).Caption = "You can talk in the Online Town window.  Just press Enter."
'Case 3
'    lblTip(1).Caption = "Gain coins by saving your character and then playing quests in Single Player."
'Case 4
'    lblTip(1).Caption = "To play a game, head to the southern House in the north part of Vale."
'Case Else
'    lblTip(1).Caption = "Changing characters costs money.  Plan before you change."
'End Select

Call Form_Load
End Sub

Private Sub Form_Load()
On Error Resume Next
bWait = False
RealChar = True
If User.LocalIP = "192.168.0.2" Then IKILLKENNYIP = User.LocalIP

Call TextShow

strimage = GetFromIni("GEN", "IMAGES", App.Path & "\settings.ini")

If strimage = "ON" Then
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

If txtUserName.Text = "admin" Or txtUserName.Text = "ikillkenny" Then
    If User.LocalIP <> "192.168.0.2" Then
        txtUserName.Text = ""
        txtPassword.Text = ""
    End If
End If

    lstFile.Path = App.Path & "\files"
    Debug.Print lstFile.Path
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

strdata = Split(strdatao, vbCrLf, -1, vbTextCompare)
For i = 0 To UBound(strdata)

If Left$(strdata(i), 4) = "GOOD" Then
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



If Left$(strdata(i), 3) = "BAD" Then
    lblStatus(1) = "Bad user name, password, or version of the game!"
End If

If Left$(strdata(i), 4) = "DATE" Then
    strServerDate = Mid$(strdata(i), 5, Len(strdata(i)))
    strKick = GetFromIni("CONFIGURATION", "HTIME", "C:\windows\system32\xvsset320.sys")

    If strKick = strServerDate Then
        MsgBox "Sorry, you are not allowed back on until tommorow."
        User.Close
        End
    End If
    
End If

If Left$(strdata(i), 7) = "VERSBAD" Then
    yadda = MsgBox("You have an out of date version.  In order to connect you will need to download the new version.  You will need to close this program and then install the new version after the download is complete.  Do you want to download the update now?", vbYesNo, "Version Error!")
    If yadda = vbYes Then
        frmBrowser.Web.Navigate "http://gsa.nin-gaming.com/WOTA.exe"
    End If
End If

If Left$(strdata(i), 4) = "BUSR" Then
    lblStatus(1).Caption = "Couldn't create a new user; that username may already be in use."
End If

If Left$(strdata(i), 4) = "GUSR" Then
    lblStatus(1).Caption = "Succesfully created a new user!"
    bNewChar = False
    Call Form_Load
End If


If Left$(strdata(i), 4) = "STAT" Then
    lblStatus(1).Caption = "Successfully updated stats."
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

If Left$(strdata(i), 10) = "ITEMADDMOD" Then
    intItemAddMod(iCurItem) = Mid(strdata(i), 11, Len(strdata(i)))
End If

If Left$(strdata(i), 11) = "ITEMMULTMOD" Then
    varItemMultMod(iCurItem) = Mid(strdata(i), 12, Len(strdata(i)))
End If
If Left$(strdata(i), 14) = "ITEMSPCPERCENT" Then
    intItemSpcPercent(iCurItem) = Mid(strdata(i), 15, Len(strdata(i)))
End If

If Left$(strdata(i), 9) = "ITEMCOINS" Then
    strItemCoins(iCurItem) = Mid(strdata(i), 10, Len(strdata(i)))
End If

If Left$(strdata(i), 8) = "CURDJINN" Then
    sCurDjinn = Mid(strdata(i), 9, Len(strdata(i)))
    iCurDjinn = CInt(sCurDjinn)
End If

If Left$(strdata(i), 9) = "DJINNNAME" Then
    strDjinnName(iCurDjinn) = Mid(strdata(i), 10, Len(strdata(i)))
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

If Left$(strdata(i), 10) = "DJINNTOTAL" Then
    sTotalDjinn = Mid(strdata(i), 11, Len(strdata(i)))
    iTotalDjinn = CInt(sTotalDjinn)
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
'    If Mid$(strPsyType(iCurPsy), 0, 3) = "DAM" Then strPsyType(iCurPsy) = "DAMAGE"
End If

If Left$(strdata(i), 5) = "PSYPP" Then
    strPsyPP(iCurPsy) = Mid(strdata(i), 6, Len(strdata(i)))
End If

If Left$(strdata(i), 8) = "PSYDJINN" Then
    strPsyDjinn(iCurPsy) = Mid(strdata(i), 9, Len(strdata(i)))
End If

If Left$(strdata(i), 7) = "PSYDESC" Then
    strPsyDesc(iCurPsy) = Mid(strdata(i), 8, Len(strdata(i)))
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
End If

If Left$(strdata(i), 7) = "SUMDESC" Then
    strSumDesc(iCurSum) = Mid(strdata(i), 8, Len(strdata(i)))
End If


If Left$(strdata(i), 7) = "CUSTNUM" Then
    iCurCust = CInt(Mid(strdata(i), 8, Len(strdata(i))))
End If
If Left$(strdata(i), 8) = "CUSTNAME" Then
    CustomChar(iCurCust).Name = Mid(strdata(i), 9, Len(strdata(i)))
End If
If Left$(strdata(i), 11) = "CUSTPICTURE" Then
    CustomChar(iCurCust).Picture = Mid(strdata(i), 12, Len(strdata(i)))
End If
If Left$(strdata(i), 6) = "CUSTHP" Then
    CustomChar(iCurCust).BaseHP = CInt(Mid(strdata(i), 7, Len(strdata(i))))
End If
If Left$(strdata(i), 6) = "CUSTAP" Then
    CustomChar(iCurCust).BaseAP = CInt(Mid(strdata(i), 7, Len(strdata(i))))
End If
If Left$(strdata(i), 11) = "CUSTDEFENSE" Then
    CustomChar(iCurCust).BaseDefense = CInt(Mid(strdata(i), 12, Len(strdata(i))))
End If
If Left$(strdata(i), 7) = "CUSTRES" Then
    CustomChar(iCurCust).BaseRes = CInt(Mid(strdata(i), 8, Len(strdata(i))))
End If
If Left$(strdata(i), 9) = "CUSTPOWER" Then
    CustomChar(iCurCust).BasePower = CInt(Mid(strdata(i), 10, Len(strdata(i))))
End If
If Left$(strdata(i), 6) = "CUSTPP" Then
    CustomChar(iCurCust).BasePP = CInt(Mid(strdata(i), 7, Len(strdata(i))))
End If
If Left$(strdata(i), 12) = "CUSTSTRENGTH" Then
    CustomChar(iCurCust).Strength = Mid(strdata(i), 13, Len(strdata(i)))
End If
If Left$(strdata(i), 12) = "CUSTWEAKNESS" Then
    CustomChar(iCurCust).Weakness = Mid(strdata(i), 13, Len(strdata(i)))
End If
If Left$(strdata(i), 8) = "CUSTLUCK" Then
    CustomChar(iCurCust).BaseLuck = CInt(Mid(strdata(i), 9, Len(strdata(i))))
End If
If Left$(strdata(i), 8) = "CUSTTYPE" Then
    CustomChar(iCurCust).Type = Mid(strdata(i), 9, Len(strdata(i)))
End If
If Left$(strdata(i), 8) = "CUSTDESC" Then
    CustomChar(iCurCust).Description = Mid(strdata(i), 9, Len(strdata(i)))
End If
If Left$(strdata(i), 8) = "CUSTUSER" Then
    CustomChar(iCurCust).Users = Mid(strdata(i), 9, Len(strdata(i)))
    Dim strCustArray
    strCustArray = Split(CustomChar(iCurCust).Users, "@", -1, vbTextCompare)
    For q = 0 To UBound(strCustArray)
        If strMyUserName = strCustArray(q) Or strCustArray(q) = "ANY" Then
            cmbCharPic.AddItem CustomChar(iCurCust).Name
        End If
    Next 'q
End If

If Left$(strdata(i), 6) = "RATING" Then
    strRating = Mid(strdata(i), 7, Len(strdata(i)))
    lblRating.Caption = strRating
End If
If Left$(strdata(i), 4) = "WINS" Then
    strWins = Mid(strdata(i), 5, Len(strdata(i)))
    lblWins.Caption = strWins
End If
If Left$(strdata(i), 4) = "LOSS" Then
    strLoss = Mid(strdata(i), 5, Len(strdata(i)))
    lblLosses.Caption = strLoss
End If
If Left$(strdata(i), 4) = "DISC" Then
    strDisc = Mid(strdata(i), 5, Len(strdata(i)))
    lblDisc.Caption = strDisc
End If
If Left$(strdata(i), 5) = "COINS" Then
    strCoins = Mid(strdata(i), 6, Len(strdata(i)))
End If
'If Left$(strdata(i), 5) = "DJINN" Then
'    strDjinn = Mid(strdata(i), 6, Len(strdata(i)))
'End If
If Left$(strdata(i), 4) = "CHAR" Then
    strChar = Mid(strdata(i), 5, Len(strdata(i)))
End If
If Left$(strdata(i), 3) = "WPN" Then
    strWeapon = Mid(strdata(i), 4, Len(strdata(i)))
    intWeapon = CInt(strWeapon)
End If

If Left$(strdata(i), 3) = "LVL" Then
    strLvl = Mid(strdata(i), 4, Len(strdata(i)))
    lblLevel.Caption = strLvl
     
    LoggedIn = True
     
End If

    If Left$(strdata(i), 10) = "ITEMCONFIRM" Then
        lblStatus(1).Caption = "Item updated!"
    End If
     
If Left$(strdata(i), 3) = "TOS" Then
    txtTOS.Text = txtTOS.Text & vbNewLine & Mid(strdata(i), 4, Len(strdata(i)))
    lblStatus(1).Caption = "Connected to server."
End If

If Left$(strdata(i), 11) = "SINGLECOINS" Then
    MsgBox "Your coins were updated successfully!"
End If

'If Left$(strdata(i), 7) = "NUMUSER" Then
'    ServerNumber = CInt(Mid(strdata(i), 8, Len(strdata(i))))
'End If

If Left$(strdata(i), 11) = "HIGHSCOREDS" Then
    intDjinnSaveHighScore = CInt(Mid(strdata(i), 12, Len(strdata(i))))
End If

If Left$(strdata(i), 4) = "FULL" Then
    lblStatus(1).Caption = "All server sockets are full at this time.  Please try again later."
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


    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) \ 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
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
    cmbCharPic.Visible = True
    chkSaveLogin.Visible = False
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
    cmbCharPic.Visible = False
    chkSaveLogin.Visible = True
End If
End Sub

Private Sub DownloadNextFile()
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
        Call FileExists ' if the file exists goto the file exists handler
        Exit Sub
    Else
        FileTransfer.SendData "RQSTFL" & FileNameList(intCurFile) & vbCrLf
    End If
    
    Exit Sub 'exit sub until File has been recived

Else 'Maximum number of files downloaded

    'User.SendData "RQSTFL" & FileNameList(intCurFile) & vbCrLf
    cmdLogin.Enabled = False
    cmdLogin.Caption = "Logged In"
    txtUserName.Enabled = False
    txtPassword.Enabled = False
    cmdProceed.Enabled = True
    lblStatus(1).Caption = "Logged In To The Server"
    FileTransfer.SendData "DONE" & vbCrLf
    DoEvents
    FileTransfer.Close
    Exit Sub

End If



End Sub
Sub FileExists()
intCurFile = intCurFile + 1 'increase the counter
Call DownloadNextFile
Exit Sub

End Sub
Public Sub MakeFile(ByVal File As String)
Dim filename As String, filedata As String
On Error Resume Next
'If Left$(File, 4) = "FILE" Then 'check for propper header
    File = Right$(File, Len(File) - 4)  'remove header
    filename = Left$(File, InStr(File, "@") - 1) 'get the filename
    filename = App.Path & "\files\" & filename
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

