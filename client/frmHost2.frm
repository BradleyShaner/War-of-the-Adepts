VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmHost2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Host A Game"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8205
   Icon            =   "frmHost2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGamePW 
      Height          =   285
      Left            =   3840
      MaxLength       =   15
      TabIndex        =   27
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtGameName 
      Height          =   285
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   25
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Game"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Host 
      Left            =   7680
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9688
   End
   Begin VB.Frame framGen 
      BackColor       =   &H00886000&
      Caption         =   "Game Options"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   7935
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start Game"
         Height          =   255
         Left            =   6600
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkTime 
         BackColor       =   &H00886000&
         Caption         =   "Timed Match"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton opLVL 
         BackColor       =   &H00886000&
         Caption         =   "40"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton opLVL 
         BackColor       =   &H00886000&
         Caption         =   "30"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton opLVL 
         BackColor       =   &H00886000&
         Caption         =   "20"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   11
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton opLVL 
         BackColor       =   &H00886000&
         Caption         =   "10"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton opLVL 
         BackColor       =   &H00886000&
         Caption         =   "5"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.CheckBox chkLadder 
         BackColor       =   &H00886000&
         Caption         =   "Ladder Match"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblReady 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Ready"
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
         Left            =   1920
         TabIndex        =   29
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opponent Is Currently:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Level To Play At:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1710
      End
   End
   Begin VB.Frame framGen 
      BackColor       =   &H00886000&
      Caption         =   "Chat"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   7935
      Begin VB.TextBox txtChatMsg 
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
         Height          =   375
         Left            =   120
         MaxLength       =   200
         TabIndex        =   5
         Top             =   1080
         Width           =   7575
      End
      Begin VB.TextBox txtChat 
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
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.Frame framGen 
      BackColor       =   &H00886000&
      Caption         =   "Your Opponent's Stats"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   3975
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   7
         Left            =   2880
         Picture         =   "frmHost2.frx":08CA
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   6
         Left            =   1920
         Picture         =   "frmHost2.frx":0FA0
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   5
         Left            =   960
         Picture         =   "frmHost2.frx":1676
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   4
         Left            =   120
         Picture         =   "frmHost2.frx":1D4C
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblRanking 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   39
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ladder Ranking:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label lblLoss 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblWins 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRating 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Losses:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   2760
         TabIndex        =   22
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wins:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rating:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame framGen 
      BackColor       =   &H00886000&
      Caption         =   "Ikillkenny's Stats"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   3
         Left            =   3000
         Picture         =   "frmHost2.frx":2422
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   2
         Left            =   2040
         Picture         =   "frmHost2.frx":2AF8
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   1
         Left            =   1080
         Picture         =   "frmHost2.frx":31CE
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   0
         Left            =   120
         Picture         =   "frmHost2.frx":38A4
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblRanking 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ladder Ranking:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label lblLoss 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   3360
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblWins 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRating 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Losses:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wins:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rating:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   3000
      TabIndex        =   26
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   930
   End
End
Attribute VB_Name = "frmHost2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opReady As Boolean

Private Sub chkLadder_Click()
On Error Resume Next
If chkLadder.Value = 0 Then
    Host.SendData "LADDERMATCHON" & vbCrLf
Else
    Host.SendData "LADDERMATCHOFF" & vbCrLf
End If
End Sub

Private Sub chkTime_Click()
On Error Resume Next
If chkTime.Value = 0 Then
    Host.SendData "TIMEMATCHOFF"
Else
    Host.SendData "TIMEMATCHON"
End If
End Sub

Private Sub cmdCreate_Click()
On Error Resume Next
If cmdCreate.Caption = "&Create Game" Then
    If Host.State <> sckClosed Then Host.Close
    DoEvents
    Host.Listen
    frmChat.Chat.SendData "METXT" & strMyUserName & " has hosted a game." & vbCrLf
    cmdCreate.Caption = "&Cancel Game"
ElseIf cmdCreate.Caption = "&Cancel Game" Then
    Host.Close
    cmdCreate.Caption = "&Create Game"
End If
End Sub

Private Sub cmdStart_Click()
On Error Resume Next
Call SendBattleData("STARTGAME")
frmBattle2.Show
Me.Hide
End Sub

Private Sub Form_Load()
Host.Close
opReady = False
framGen(3).Enabled = False
framGen(2).Enabled = False
cmdCreate.Enabled = True
txtGameName.Enabled = True
txtGamePW.Enabled = True
End Sub

Private Sub Host_Connect()
On Error Resume Next
txtChat.Text = txtChat.Text & vbNewLine & "Client connected!"
Me.cmdCreate.Enabled = False
Me.txtGameName.Enabled = False
Me.txtGamePW.Enabled = False
framGen(2).Enabled = True
framGen(3).Enabled = True
End Sub

Private Sub Host_DataArrival(ByVal bytesTotal As Long)
On Error GoTo err
Dim strRawData As String

Dim strTime As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")

Host.GetData strRawData
strData = Split(strRawData, vbCrLf, -1, vbTextCompare) 'Split packets between @'s

For i = 0 To UBound(strData)
    If Left$(strData(i), 8) = "USERNAME" Then
        strFoeUserName = Mid$(strData(i), 9, Len(strData(i)))
        framGen(1).Caption = strFoeUserName & "'s stats."
        Host.SendData "USERNAME" & strMyUserName & vbCrLf
        Host.SendData "RATING" & strMyRating & vbCrLf
        Host.SendData "WINS" & strMyWins & vbCrLf
        Host.SendData "LOSS" & strMyLoss & vbCrLf
        Host.SendData "RANKING" & strMyRanking & vbCrLf
    End If
    If Left$(strData(i), 6) = "RATING" Then
        strFoeRating = Mid$(strData(i), 7, Len(strData(i)))
        lblRating(1).Caption = strFoeRating
    End If
    If Left$(strData(i), 4) = "WINS" Then
        strFoeWins = Mid$(strData(i), 5, Len(strData(i)))
        lblWins(1).Caption = strFoeWins
    End If
    If Left$(strData(i), 4) = "LOSS" Then
        strFoeLoss = Mid$(strData(i), 5, Len(strData(i)))
        lblLoss(1).Caption = strFoeLoss
    End If
    If Left$(strData(i), 6) = "RANKING" Then
        strFoeRanking = Mid$(strData(i), 7, Len(strData(i)))
        lblRanking(1).Caption = strFoeRanking
    End If
    If Left$(strData(i), 9) = "PARTYNAME" Then
        Dim intTempChar As Long
        intTempChar = CLng(Mid$(strData(i), 10, 1))
        WOTAChar(intTempChar + 4).Name = Mid$(strData(i), 11, Len(strData(i)))
        imgParty(intTempChar + 4).Picture = LoadPicture(App.Path & "\BattleImages\" & WOTAChar(intTempChar + 4).Name & ".gif")
    End If
    If Left$(strData(i), 8) = "WAITCHAT" Then
        Dim strTempWaitChat As String
        strTempWaitChat = Mid$(strData(i), 9, Len(strData(i)))
        txtChat.Text = txtChat.Text & vbNewLine & strTempWaitChat
    End If
Next 'i

Exit Sub
err:
Exit Sub
End Sub

Private Sub opLVL_Click(Index As Integer)
On Error Resume Next
Host.SendData "BATTLELVL" & opLVL(Index).Caption & vbCrLf
End Sub

Private Sub txtChat_Change()
Call AutoScrollTxt(txtChat)
End Sub

Private Sub txtChatMsg_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    Host.SendData "WAITCHAT" & strMyUserName & ": " & txtChatMsg.Text & vbCrLf
    txtChat.Text = txtChat.Text & vbNewLine & strMyUserName & ": " & txtChatMsg.Text
    txtChatMsg.Text = ""
End If

End Sub
